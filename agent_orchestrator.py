import os
import json
import shutil
import asyncio
import requests
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv
from rapidfuzz import fuzz
from google.antigravity import Agent, LocalAgentConfig

# --- API KEY CONFIGURATION ---
# Set your Gemini API key in the .env file (GEMINI_API_KEY=...) or enter it once in the UI.
HARDCODED_API_KEY = ""

# Recommended modern models from Google Gemini API (priority queue for automatic fallback)
SUPPORTED_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.1-flash-lite",
    "gemini-3.6-flash",
    "gemini-3.7-flash",
    "gemini-flash-latest"
]
DEFAULT_MODEL = "gemini-3.8-flash"


HISTORY_FILE = "trivia_history.json"
CACHE_DIR = "cache"
KEY_FILE = "gemini_api_key.txt"
ENV_FILE = ".env"

load_dotenv()

def safe_print(text: str):
    try:
        print(text)
    except Exception:
        try:
            print(str(text).encode("ascii", errors="replace").decode("ascii"))
        except Exception:
            pass

def get_saved_api_key() -> str:
    """Retrieve saved Gemini API key from hardcoded setting, .env, key file, or environment."""
    if HARDCODED_API_KEY and HARDCODED_API_KEY.strip():
        return HARDCODED_API_KEY.strip()
    env_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if env_key:
        return env_key
    if os.path.exists(KEY_FILE):
        try:
            with open(KEY_FILE, "r", encoding="utf-8") as f:
                k = f.read().strip()
                if k:
                    return k
        except Exception:
            pass
    if os.path.exists(ENV_FILE):
        try:
            with open(ENV_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    if line.strip().startswith("GEMINI_API_KEY="):
                        k = line.strip().split("=", 1)[1].strip().strip('"\'')
                        if k:
                            return k
        except Exception:
            pass
    return ""

def persist_api_key(api_key: str):
    """Persist the API key locally to .env and gemini_api_key.txt so it never needs to be re-entered."""
    if not api_key or not api_key.strip():
        return
    api_key = api_key.strip()
    os.environ["GEMINI_API_KEY"] = api_key
    try:
        with open(KEY_FILE, "w", encoding="utf-8") as f:
            f.write(api_key)
    except Exception as e:
        print(f"Error saving to {KEY_FILE}: {e}")
    try:
        env_lines = []
        if os.path.exists(ENV_FILE):
            with open(ENV_FILE, "r", encoding="utf-8") as f:
                env_lines = [l for l in f.readlines() if not l.strip().startswith("GEMINI_API_KEY=")]
        env_lines.append(f"GEMINI_API_KEY={api_key}\n")
        with open(ENV_FILE, "w", encoding="utf-8") as f:
            f.writelines(env_lines)
    except Exception as e:
        print(f"Error saving to {ENV_FILE}: {e}")

def load_history(filepath: str = HISTORY_FILE) -> list:
    """Load historical questions from trivia_history.json."""
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception as e:
            print(f"Error loading history file {filepath}: {e}")
    return []

def save_history(history: list, filepath: str = HISTORY_FILE):
    """Save updated history to trivia_history.json."""
    try:
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Error saving history file {filepath}: {e}")

def load_round_3_from_csv(csv_path: str = "questions.csv") -> list:
    """
    Load Round 3 questions from existing questions.csv and build candidate items with image search queries.
    Uses song names and artists from answers to construct optimal search queries for images.
    """
    items = []
    if os.path.exists(csv_path):
        try:
            df = pd.read_csv(csv_path)
            if "Round" in df.columns:
                r3_df = df[df["Round"] == 3]
                for idx, (_, row) in enumerate(r3_df.iterrows(), 1):
                    q_text = str(row.get("Question", ""))
                    ans_text = str(row.get("Answer", ""))
                    r_name = str(row.get("Round_Name", "Ghicește Melodia"))
                    pic_file = str(row.get("Picture", f"{idx}.jpg"))

                    parts = ans_text.split("-", 1) if "-" in ans_text else [ans_text, ""]
                    song = parts[0].strip()
                    artist = parts[1].strip() if len(parts) > 1 else ""
                    query = f"{artist} {song}" if artist else song

                    items.append({
                        "round_nr": 3,
                        "round_name": r_name,
                        "question": q_text,
                        "answer": ans_text,
                        "image_query": query,
                        "picture": pic_file
                    })
        except Exception as e:
            safe_print(f"Error loading Round 3 from {csv_path}: {e}")
    return items


def is_duplicate(cand: dict, history: list, threshold: float = 80.0) -> bool:
    """
    Check if a candidate question/answer matches any historical item using rapidfuzz token_set_ratio.
    Returns True if ratio > threshold.
    For Round 2 (question == 'Poza'), compares answers.
    """
    q_cand = str(cand.get("question", "")).strip().lower()
    a_cand = str(cand.get("answer", "")).strip().lower()
    if not q_cand and not a_cand:
        return False
        
    for h in history:
        h_q = str(h.get("question", "")).strip().lower()
        h_a = str(h.get("answer", "")).strip().lower()
        
        # Round 2 case: question is "Poza", compare answers
        if q_cand == "poza":
            if h_a and fuzz.token_set_ratio(a_cand, h_a) > threshold:
                return True
        else:
            if h_q and fuzz.token_set_ratio(q_cand, h_q) > threshold:
                return True
            # Secondary check: if answer matches very closely and question matches somewhat
            if h_a and fuzz.token_set_ratio(a_cand, h_a) > 85.0 and fuzz.token_set_ratio(q_cand, h_q) > 60.0:
                return True
    return False

def extract_json(text: str):
    """Extract JSON object or array from model response string."""
    text = text.strip()
    if text.startswith("```json"):
        text = text[7:]
    if text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    text = text.strip()
    return json.loads(text)

def fetch_and_cache_image(query: str, save_path: str, index: int = 0) -> bool:
    """
    Search DuckDuckGo images for `query` and download the result at `index` to `save_path`.
    Returns True if successfully downloaded.
    """
    if not query or not query.strip():
        return False
    try:
        try:
            from ddgs import DDGS
        except ImportError:
            from duckduckgo_search import DDGS
            
        ddgs = DDGS()
        results = list(ddgs.images(query, max_results=max(15, index + 6)))
        if not results:
            return False
            
        start_idx = min(index, len(results) - 1)
        for idx in range(start_idx, len(results)):
            img_url = results[idx].get("image") or results[idx].get("thumbnail")
            if not img_url:
                continue
            try:
                resp = requests.get(
                    img_url,
                    headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
                    timeout=6
                )
                if resp.status_code == 200 and len(resp.content) > 1000:
                    os.makedirs(os.path.dirname(save_path), exist_ok=True)
                    with open(save_path, "wb") as f:
                        f.write(resp.content)
                    return True
            except Exception:
                continue
    except Exception as e:
        print(f"Error fetching image for query '{query}': {e}")
    return False


async def _call_llm_with_resilience(
    api_key: str,
    prompt: str,
    model_name: str = DEFAULT_MODEL,
    status_callback = None
) -> str:
    """
    Invokes LLM with automatic model failover.
    When a quota limit (429 / Resource Exhausted) or unavailability (503 / 404) is hit,
    it automatically and immediately switches to the next model in the priority queue.
    """
    models_to_try = [model_name] if model_name else []
    for m in SUPPORTED_MODELS:
        if m not in models_to_try:
            models_to_try.append(m)

    last_error = None

    for m in models_to_try:
        config = LocalAgentConfig(
            api_key=api_key if api_key else None,
            model=m,
            system_instructions="You are an expert Romanian trivia master and question generator. Always return valid, well-structured JSON only without markdown conversational wrapper."
        )

        try:
            if status_callback:
                status_callback(f"🤖 Se generează cu modelul {m}...")
            safe_print(f"Attempting generation with model: {m}")
            async with Agent(config) as agent:
                resp = await agent.chat(prompt)
                txt = await resp.text()
                if txt and txt.strip():
                    return txt
        except Exception as e:
            last_error = e
            err_str = str(e).lower()
            safe_print(f"Model {m} error: {e}")
            if "429" in err_str or "quota" in err_str or "resource_exhausted" in err_str:
                msg = f"⚠️ Cota atinsă pe modelul '{m}'. Se comută automat pe următorul model..."
                safe_print(f"Model {m} hit quota, switching to next model...")
                if status_callback:
                    status_callback(msg)
                continue
            elif "503" in err_str or "unavailable" in err_str or "404" in err_str or "not found" in err_str:
                msg = f"⚠️ Modelul '{m}' este indisponibil. Se comută automat pe următorul model..."
                safe_print(f"Model {m} unavailable, switching to next model...")
                if status_callback:
                    status_callback(msg)
                continue
            else:
                # Any other error: move to next model
                continue



    # Fallback to direct google.genai client if needed
    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        for m in models_to_try:
            try:
                if status_callback:
                    status_callback(f"🤖 Conexiune directă cu {m}...")
                res = await client.aio.models.generate_content(
                    model=m,
                    contents=prompt
                )
                if res.text and res.text.strip():
                    return res.text
            except Exception as e_genai:
                last_error = e_genai
                continue
    except Exception:
        pass

    raise Exception(f"Nu s-a putut genera conținutul. Verificați cheia API.\nEroare: {last_error}")


async def run_full_generation_pipeline(
    api_key: str,
    round_configs: list,
    progress_callback=None,
    model_name: str = DEFAULT_MODEL
) -> dict:
    """
    Batched generation pipeline: generates all rounds in 1 single optimized API call to avoid 5 RPM rate limits.
    """
    if not api_key or not api_key.strip():
        api_key = get_saved_api_key()
    if api_key:
        persist_api_key(api_key)

    history = load_history()
    
    # Extract thematic names
    theme4 = next((cfg["round_name"] for cfg in round_configs if cfg["round_nr"] == 4), "Tematica 1")
    theme5 = next((cfg["round_name"] for cfg in round_configs if cfg["round_nr"] == 5), "Tematica 2")

    target_counts = {1: 15, 2: 15, 4: 15, 5: 15, 6: 5}

    prompt = f"""Ești un maestru de trivia pentru un eveniment live din Transilvania, România.
Generează un set complet de întrebări de trivia captivante și de cultură generală în limba Română, formatat ca JSON.

ESTE ESENȚIAL: Respectă cu strictețe numărul minim de întrebări cerut pentru fiecare rundă (generează o listă completă, fără a te opri mai devreme):
1. Runda 1: "Cultură Generală" -> EXACT 20 întrebări diverse (istorie, știință, geografie, literatură, artă).
2. Runda 2: "Ce/Cine se află în această poză?" -> EXACT 20 elemente vizuale diferite.
   IMPORTANT PENTRU RUNDA 2:
   - Câmpul "question" TREBUIE SĂ FIE ÎNTOTDEAUNA exact cuvântul "Poza".
   - "answer": Numele / entitatea corectă de ghicit.
   - "image_query_a": Căutare scurtă în engleză/română pentru indiciul imagine A.
   - "image_query_b": Căutare scurtă pentru indiciul imagine B.
   - "image_query_c": Căutare scurtă pentru indiciul imagine C.
3. Runda 4: "{theme4}" -> EXACT 20 întrebări interesante pe tema "{theme4}".
   - "question": Textul întrebării.
   - "answer": Răspunsul corect.
   - "image_query": Căutare imagine relevantă.
4. Runda 5: "{theme5}" -> EXACT 20 întrebări interesante pe tema "{theme5}".
   - "question": Textul întrebării.
   - "answer": Răspunsul corect.
   - "image_query": Căutare imagine relevantă.
5. Runda 6: "Pariul" (Wager) -> EXACT 8 întrebări capcană cu 4 variante de răspuns (A, B, C, D) unde răspunsul intuitiv este greșit.
   - "question": Întrebarea completă incluzând opțiunile A), B), C), D).
   - "answer": Opțiunea corectă și explicația scurtă (ex: "B) Mauna Kea (10210 m)").
   - "image_query": Căutare imagine de fundal.

FORMATUL JSON EXACT CERUT (returnează DOAR JSON valid):
{{
  "round_1": [
    {{"question": "...", "answer": "...", "image_query": "..."}}
  ],
  "round_2": [
    {{"question": "Poza", "answer": "...", "image_query_a": "...", "image_query_b": "...", "image_query_c": "..."}}
  ],
  "round_4": [
    {{"question": "...", "answer": "...", "image_query": "..."}}
  ],
  "round_5": [
    {{"question": "...", "answer": "...", "image_query": "..."}}
  ],
  "round_6": [
    {{"question": "...", "answer": "...", "image_query": "..."}}
  ]
}}
"""

    status_cb = (lambda msg: progress_callback(1, 4, msg)) if progress_callback else None
    if progress_callback:
        progress_callback(1, 4, f"Se generează setul complet de întrebări cu {model_name}...")

    raw_text = await _call_llm_with_resilience(api_key, prompt, model_name=model_name, status_callback=status_cb)
    data = extract_json(raw_text)

    results = {}
    r_names_map = {
        1: "Cultură Generală",
        2: "Ce/Cine se află în această poză?",
        4: theme4,
        5: theme5,
        6: "Pariul"
    }

    for r_nr in [1, 2, 4, 5, 6]:
        key = f"round_{r_nr}"
        raw_items = data.get(key, [])
        valid_items = []
        for it in raw_items:
            if r_nr == 2:
                it["question"] = "Poza"
            it["round_nr"] = r_nr
            it["round_name"] = r_names_map[r_nr]
            
            # Deduplication
            if not is_duplicate(it, history) and not is_duplicate(it, valid_items):
                valid_items.append(it)

        results[r_nr] = valid_items

    # Top-up check: If any round has fewer than target_counts, run a targeted top-up
    missing_rounds = {r: target_counts[r] - len(results[r]) for r in target_counts if len(results[r]) < target_counts[r]}
    if missing_rounds:
        top_up_prompts = []
        for r_nr, needed in missing_rounds.items():
            top_up_prompts.append(f"- Runda {r_nr} ({r_names_map[r_nr]}): {needed + 3} întrebări adiționale.")
        
        topup_prompt = f"""Generează întrebări suplimentare de trivia în limba Română, în format JSON, pentru a completa numărul necesar:
{chr(10).join(top_up_prompts)}

Pentru Runda 2, "question" trebuie să fie "Poza", cu "image_query_a", "image_query_b" și "image_query_c".
Pentru Runda 6, întrebări cu opțiuni A, B, C, D.

Returnează DOAR JSON cu cheile corespunzătoare: {list(f'round_{r}' for r in missing_rounds.keys())}
"""
        try:
            if progress_callback:
                progress_callback(2, 4, "Se completează întrebările lipsă pentru a atinge numărul cerut...")
            raw_topup = await _call_llm_with_resilience(api_key, topup_prompt, model_name=model_name, status_callback=status_cb)
            topup_data = extract_json(raw_topup)
            for r_nr in missing_rounds:
                key = f"round_{r_nr}"
                for it in topup_data.get(key, []):
                    if r_nr == 2:
                        it["question"] = "Poza"
                    it["round_nr"] = r_nr
                    it["round_name"] = r_names_map[r_nr]
                    if not is_duplicate(it, history) and not is_duplicate(it, results[r_nr]):
                        results[r_nr].append(it)
        except Exception as e:
            print(f"Top-up warning: {e}")

    # Final guaranteed sizing
    for r_nr, count in target_counts.items():
        results[r_nr] = results[r_nr][:count]

    # Load Round 3 from questions.csv (take song names & artists to generate images)
    results[3] = load_round_3_from_csv("questions.csv")

    return results



async def generate_more_candidates_for_round(
    api_key: str,
    round_nr: int,
    round_name: str,
    count: int = 10,
    existing_candidates: list = None,
    model_name: str = DEFAULT_MODEL,
    status_callback = None
) -> list:
    """
    Generate additional candidates for a specific round, ensuring no duplicates with existing items or history.
    """
    if not api_key or not api_key.strip():
        api_key = get_saved_api_key()
    if api_key:
        persist_api_key(api_key)

    history = load_history()
    existing_candidates = existing_candidates or []

    oversample = count + 4

    if round_nr == 1:
        instructions = f"""Generează EXACT {oversample} întrebări NOI, diverse și captivante de Cultură Generală (istorie, știință, geografie, literatură, artă) în limba Română."""
        format_example = '[{"question": "...", "answer": "...", "image_query": "..."}]'
    elif round_nr == 2:
        instructions = f"""Generează EXACT {oversample} elemente vizuale NOI (personalități, filme, locuri, personaje, obiecte celebre) în limba Română.
IMPORTANT: Câmpul "question" TREBUIE SĂ FIE ÎNTOTDEAUNA exact "Poza". "answer" este entitatea de ghicit. Furnizează "image_query_a", "image_query_b" și "image_query_c"."""
        format_example = '[{"question": "Poza", "answer": "...", "image_query_a": "...", "image_query_b": "...", "image_query_c": "..."}]'
    elif round_nr in [4, 5]:
        instructions = f"""Generează EXACT {oversample} întrebări NOI și interesante pe tema "{round_name}" în limba Română."""
        format_example = '[{"question": "...", "answer": "...", "image_query": "..."}]'
    elif round_nr == 6:
        instructions = f"""Generează EXACT {oversample} întrebări capcană NOI cu 4 variante de răspuns (A, B, C, D) unde răspunsul intuitiv este greșit în limba Română."""
        format_example = '[{"question": "...", "answer": "...", "image_query": "..."}]'
    else:
        instructions = f"""Generează EXACT {oversample} întrebări NOI de trivia în limba Română."""
        format_example = '[{"question": "...", "answer": "...", "image_query": "..."}]'

    prompt = f"""Ești un maestru de trivia pentru un eveniment live din Transilvania, România.
{instructions}

IMPORTANT: Întrebările trebuie să fie complet noi și unice, fără a se repeta.
Returnează DOAR un array JSON valid:
{format_example}
"""

    raw_text = await _call_llm_with_resilience(api_key, prompt, model_name=model_name, status_callback=status_callback)

    try:
        raw_items = extract_json(raw_text)
        if isinstance(raw_items, dict):
            # If wrapped in a dictionary, take the first list value
            for v in raw_items.values():
                if isinstance(v, list):
                    raw_items = v
                    break
    except Exception as e:
        print(f"Error parsing more candidates JSON: {e}")
        return []

    valid_new = []
    for it in raw_items:
        if round_nr == 2:
            it["question"] = "Poza"
        it["round_nr"] = round_nr
        it["round_name"] = round_name

        if not is_duplicate(it, history) and not is_duplicate(it, existing_candidates) and not is_duplicate(it, valid_new):
            valid_new.append(it)

    return valid_new[:count]


def download_additional_candidate_images(round_nr: int, new_candidates: list, start_index: int = 0, cache_dir: str = CACHE_DIR):
    """
    Download images for newly added candidates, indexing from start_index.
    """
    round_cache = os.path.join(cache_dir, f"round_{round_nr}")
    os.makedirs(round_cache, exist_ok=True)

    for idx_offset, cand in enumerate(new_candidates):
        idx = start_index + idx_offset
        if round_nr == 2:
            q_a = cand.get("image_query_a") or f"{cand.get('answer')} clue 1"
            q_b = cand.get("image_query_b") or f"{cand.get('answer')} clue 2"
            q_c = cand.get("image_query_c") or f"{cand.get('answer')} photo"

            path_a = os.path.join(round_cache, f"cand_{idx}_A.jpg")
            path_b = os.path.join(round_cache, f"cand_{idx}_B.jpg")
            path_c = os.path.join(round_cache, f"cand_{idx}_C.jpg")

            fetch_and_cache_image(q_a, path_a, index=0)
            fetch_and_cache_image(q_b, path_b, index=0)
            fetch_and_cache_image(q_c, path_c, index=0)

            cand["cached_img_a"] = path_a
            cand["cached_img_b"] = path_b
            cand["cached_img_c"] = path_c
            cand["img_index_a"] = 0
            cand["img_index_b"] = 0
            cand["img_index_c"] = 0
        else:
            q = cand.get("image_query") or cand.get("answer") or cand.get("question")
            path = os.path.join(round_cache, f"cand_{idx}.jpg")
            fetch_and_cache_image(q, path, index=0)
            cand["cached_img"] = path
            cand["img_index"] = 0


def download_candidate_images(candidates_by_round: dict, cache_dir: str = CACHE_DIR, progress_callback=None):

    """
    Download candidate images into cache directory.
    For Round 2, downloads triplets of distinct images (A, B, C).
    """
    if os.path.exists(cache_dir):
        shutil.rmtree(cache_dir, ignore_errors=True)
    os.makedirs(cache_dir, exist_ok=True)
    
    total_items = sum(len(cands) for cands in candidates_by_round.values())
    done_items = 0
    
    for round_nr, cands in candidates_by_round.items():
        round_cache = os.path.join(cache_dir, f"round_{round_nr}")
        os.makedirs(round_cache, exist_ok=True)
        
        for idx, cand in enumerate(cands):
            if round_nr == 2:
                q_a = cand.get("image_query_a") or f"{cand.get('answer')} clue 1"
                q_b = cand.get("image_query_b") or f"{cand.get('answer')} clue 2"
                q_c = cand.get("image_query_c") or f"{cand.get('answer')} photo"
                
                path_a = os.path.join(round_cache, f"cand_{idx}_A.jpg")
                path_b = os.path.join(round_cache, f"cand_{idx}_B.jpg")
                path_c = os.path.join(round_cache, f"cand_{idx}_C.jpg")
                
                fetch_and_cache_image(q_a, path_a, index=0)
                fetch_and_cache_image(q_b, path_b, index=0)
                fetch_and_cache_image(q_c, path_c, index=0)
                
                cand["cached_img_a"] = path_a
                cand["cached_img_b"] = path_b
                cand["cached_img_c"] = path_c
                cand["img_index_a"] = 0
                cand["img_index_b"] = 0
                cand["img_index_c"] = 0
            else:
                q = cand.get("image_query") or cand.get("answer") or cand.get("question")
                path = os.path.join(round_cache, f"cand_{idx}.jpg")
                fetch_and_cache_image(q, path, index=0)
                cand["cached_img"] = path
                cand["img_index"] = 0
                
            done_items += 1
            if progress_callback:
                progress_callback(done_items, total_items, f"Descărcare imagine Runda {round_nr} (#{idx+1}/{len(cands)})")

def regenerate_candidate_image(cand: dict, round_nr: int, image_slot: str = "main", custom_query: str = None) -> bool:
    """
    Re-fetch DuckDuckGo image for a candidate using next search result index, or using a custom query if specified.
    `image_slot`: 'main' (for regular rounds), 'A', 'B', or 'C' (for Round 2).
    `custom_query`: Optional updated search query string. If provided, updates cand and resets index to 0.
    """
    if round_nr == 2:
        if image_slot == "A":
            if custom_query and custom_query.strip():
                cand["image_query_a"] = custom_query.strip()
                cand["img_index_a"] = 0
            else:
                cand["img_index_a"] = cand.get("img_index_a", 0) + 1
            idx = cand.get("img_index_a", 0)
            q = cand.get("image_query_a") or cand.get("answer")
            save_path = cand.get("cached_img_a")
            return fetch_and_cache_image(q, save_path, index=idx)
        elif image_slot == "B":
            if custom_query and custom_query.strip():
                cand["image_query_b"] = custom_query.strip()
                cand["img_index_b"] = 0
            else:
                cand["img_index_b"] = cand.get("img_index_b", 0) + 1
            idx = cand.get("img_index_b", 0)
            q = cand.get("image_query_b") or cand.get("answer")
            save_path = cand.get("cached_img_b")
            return fetch_and_cache_image(q, save_path, index=idx)
        elif image_slot == "C":
            if custom_query and custom_query.strip():
                cand["image_query_c"] = custom_query.strip()
                cand["img_index_c"] = 0
            else:
                cand["img_index_c"] = cand.get("img_index_c", 0) + 1
            idx = cand.get("img_index_c", 0)
            q = cand.get("image_query_c") or cand.get("answer")
            save_path = cand.get("cached_img_c")
            return fetch_and_cache_image(q, save_path, index=idx)
    else:
        if custom_query and custom_query.strip():
            cand["image_query"] = custom_query.strip()
            cand["img_index"] = 0
        else:
            cand["img_index"] = cand.get("img_index", 0) + 1
        idx = cand.get("img_index", 0)
        q = cand.get("image_query") or cand.get("answer") or cand.get("question")
        save_path = cand.get("cached_img")
        return fetch_and_cache_image(q, save_path, index=idx)


def confirm_and_finalize(
    selected_by_round: dict,
    unselected_by_round: dict,
    history_file: str = HISTORY_FILE,
    generator_fn = None
) -> tuple[str, int]:
    """
    Executes final confirmation:
    1. Appends unselected candidates to history JSON as 'discarded'.
    2. Appends accepted items to history JSON as 'accepted'.
    3. Moves selected images from cache/ to images/Round X/ folders.
    4. Saves final dataset to questions.csv.
    5. Triggers generate_trivia_slides(pd.read_csv('questions.csv')).
    Returns path to generated pptx and count of exported questions.
    """
    history = load_history(history_file)
    now_str = datetime.now().isoformat()
    
    # 1. Record discarded items
    for r_nr, cands in unselected_by_round.items():
        for cand in cands:
            history.append({
                "round": r_nr,
                "round_name": cand.get("round_name"),
                "question": cand.get("question"),
                "answer": cand.get("answer"),
                "status": "discarded",
                "timestamp": now_str
            })
            
    # 2. Record accepted items & build CSV rows
    existing_r3_rows = []
    if 3 not in selected_by_round and os.path.exists("questions.csv"):
        try:
            old_df = pd.read_csv("questions.csv")
            if "Round" in old_df.columns:
                r3_df = old_df[old_df["Round"] == 3]
                if not r3_df.empty:
                    existing_r3_rows = r3_df.to_dict(orient="records")
        except Exception:
            pass

    csv_rows = []
    
    for r_nr in sorted(selected_by_round.keys()):
        cands = selected_by_round[r_nr]
        r_name = cands[0].get("round_name") if cands else f"Round {r_nr}"
        
        target_img_dir = "images/Wager" if r_nr == 6 else f"images/Round {r_nr}"
        os.makedirs(target_img_dir, exist_ok=True)
        
        for idx, cand in enumerate(cands, 1):
            history.append({
                "round": r_nr,
                "round_name": r_name,
                "question": cand.get("question"),
                "answer": cand.get("answer"),
                "status": "accepted",
                "timestamp": now_str
            })
            
            if r_nr == 2:
                # Move triplet of images: {idx}A, {idx}B, and {idx}C
                src_a = cand.get("cached_img_a")
                src_b = cand.get("cached_img_b")
                src_c = cand.get("cached_img_c")
                
                ext_a = os.path.splitext(src_a)[1] if src_a and os.path.splitext(src_a)[1] else ".jpg"
                ext_b = os.path.splitext(src_b)[1] if src_b and os.path.splitext(src_b)[1] else ".jpg"
                ext_c = os.path.splitext(src_c)[1] if src_c and os.path.splitext(src_c)[1] else ".jpg"
                
                dest_a = os.path.join(target_img_dir, f"{idx}A{ext_a}")
                dest_b = os.path.join(target_img_dir, f"{idx}B{ext_b}")
                dest_c = os.path.join(target_img_dir, f"{idx}C{ext_c}")
                
                if src_a and os.path.exists(src_a):
                    shutil.copy2(src_a, dest_a)
                if src_b and os.path.exists(src_b):
                    shutil.copy2(src_b, dest_b)
                if src_c and os.path.exists(src_c):
                    shutil.copy2(src_c, dest_c)
                    
                picture_col = f"{idx}B{ext_b}"
                question_col = "Poza"
            elif r_nr == 6:
                # Wager image
                src = cand.get("cached_img")
                ext = os.path.splitext(src)[1] if src and os.path.splitext(src)[1] else ".jpeg"
                dest = os.path.join(target_img_dir, f"Wager{ext}")
                if src and os.path.exists(src):
                    shutil.copy2(src, dest)
                picture_col = f"Wager{ext}"
                question_col = cand.get("question")
            else:
                # Rounds 1, 3, 4, 5
                src = cand.get("cached_img")
                ext = os.path.splitext(src)[1] if src and os.path.splitext(src)[1] else ".jpg"
                dest = os.path.join(target_img_dir, f"{idx}{ext}")
                if src and os.path.exists(src):
                    shutil.copy2(src, dest)
                picture_col = f"{idx}{ext}"
                question_col = cand.get("question")
                
            cand_round_name = cand.get("round_name") or r_name
            csv_rows.append({
                "Round": r_nr,
                "Round_Name": cand_round_name,
                "Question": question_col,
                "Answer": cand.get("answer"),
                "Picture": picture_col
            })
            
        if r_nr == 2 and existing_r3_rows:
            csv_rows.extend(existing_r3_rows)
            
    # Save history
    save_history(history, history_file)
    
    # Save questions.csv
    df_out = pd.DataFrame(csv_rows)
    df_out.to_csv("questions.csv", index=False, encoding="utf-8")
    
    # Execute PPT generation
    if generator_fn is not None:
        pptx_path = generator_fn(pd.read_csv("questions.csv"))
    else:
        import sys
        if "app" in sys.modules:
            pptx_path = sys.modules["app"].generate_trivia_slides(pd.read_csv("questions.csv"))
        else:
            from app import generate_trivia_slides
            pptx_path = generate_trivia_slides(pd.read_csv("questions.csv"))
    
    return pptx_path, len(csv_rows)

