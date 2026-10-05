import sys
import os
import asyncio

# --- 1. PYTHON 3.13 & FFmpeg PATH STRIKE ---
current_dir = os.path.dirname(os.path.abspath(__file__))
ffmpeg_dir = os.path.join(current_dir, "ffmpeg")
os.environ["PATH"] += os.pathsep + ffmpeg_dir

try:
    import audioop
except ImportError:
    try:
        import audioop_lts as audioop

        sys.modules["audioop"] = audioop
    except ImportError:
        pass

import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from PIL import Image
from pydub import AudioSegment
import glob
import io
import re

# Anchor pydub to the folder
ffmpeg_exe = os.path.join(ffmpeg_dir, "ffmpeg.exe")
if os.path.exists(ffmpeg_exe):
    AudioSegment.converter = ffmpeg_exe
    AudioSegment.ffprobe = os.path.join(ffmpeg_dir, "ffprobe.exe")
else:
    st.error(f"❌ FFmpeg NOT FOUND at: {ffmpeg_exe}")

# --- PAGE CONFIG ---
st.set_page_config(page_title="Transylvania Trivia Command", layout="wide", page_icon="🧛")

# --- FOLDER SETUP ---
GRADING_FOLDER = "Grading Sheets"
MASTER_FILE_PATH = os.path.join(GRADING_FOLDER, "Grading Sheet Overall.xlsx")
if not os.path.exists(GRADING_FOLDER):
    os.makedirs(GRADING_FOLDER)


# --- HELPERS ---
def get_ordinal(n):
    if 11 <= (n % 100) <= 13: return f"{n}th"
    return f"{n}" + {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')


def get_dynamic_font_size(text, is_qa=False):
    length = len(str(text))
    if is_qa: return 64 if length < 15 else (52 if length < 30 else 42)
    return 58 if length < 30 else (48 if length < 70 else 38)


# --- 1. THE ADVANCED SCORING ENGINE ---
def run_calculation(df):
    if df.empty: return df
    numeric_cols = ['R1', 'R2', 'R3', 'R4', 'R5', 'Joker Pct', 'Pariu (±)']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.00)

    def row_math(row):
        rounds = ["R1", "R2", "R3", "R4", "R5"]
        score_sum = 0.00
        j_choice = str(row.get('Joker', 'Niciuna'))
        j_pct = float(row.get('Joker Pct', 0.00))
        for r in rounds:
            val = float(row[r])
            if j_choice == r:
                if val >= j_pct:
                    score_sum += val + j_pct
                else:
                    score_sum += val
            else:
                score_sum += val
        return score_sum + float(row.get('Pariu (±)', 0.00))

    df['Total'] = df.apply(row_math, axis=1)
    return df


def manage_scores():
    st.header("🏆 Trivia Scoring Matrix")
    COLS = ['Echipă', 'R1', 'R2', 'R3', 'R4', 'R5', 'Joker', 'Joker Pct', 'Pariu (±)', 'Total']
    if 'teams' not in st.session_state:
        st.session_state.teams = pd.DataFrame(columns=COLS)

    with st.expander("➕ Adaugă Echipă Nouă", expanded=True):
        t_name = st.text_input("Nume Echipă", key="new_team_name_input")
        if st.button("Înregistrează Echipa"):
            if t_name:
                new_t = pd.DataFrame(
                    [{'Echipă': t_name, 'R1': 0.00, 'R2': 0.00, 'R3': 0.00, 'R4': 0.00, 'R5': 0.00, 'Joker': "Niciuna",
                      'Joker Pct': 0.00, 'Pariu (±)': 0.00, 'Total': 0.00}])
                st.session_state.teams = pd.concat([st.session_state.teams, new_t], ignore_index=True)
                st.rerun()
    st.markdown("---")
    col_input, col_reset, col_push = st.columns([1, 1, 1])
    with col_input:
        edition = st.number_input("Ediția Trivia", min_value=1, step=1, value=1, label_visibility="collapsed")
    with col_reset:
        if st.button("🧹 Reset / Șterge Tot", use_container_width=True):
            st.session_state.teams = pd.DataFrame(columns=COLS)
            if 'reveal_step' in st.session_state: st.session_state.reveal_step = 0
            st.rerun()
    with col_push:
        if not st.session_state.teams.empty:
            if st.button(f"🚀 Push to Ed. {edition}", use_container_width=True):
                # Individual Logic and Master Sheet Logic Integrated Here
                edition_str = get_ordinal(edition)
                session_path = os.path.join(GRADING_FOLDER, f"Grading Sheet {edition_str} edition.xlsx")
                output_session = io.BytesIO()
                with pd.ExcelWriter(output_session, engine='xlsxwriter') as writer:
                    export_df = st.session_state.teams.sort_values(by='Total', ascending=False)
                    export_df.to_excel(writer, index=False, sheet_name='Final Scores')
                with open(session_path, "wb") as f:
                    f.write(output_session.getvalue())

                # 2. UPDATE MASTER OVERALL SHEET
                edition_col = f"Ediția {edition}"
                if os.path.exists(MASTER_FILE_PATH):
                    df_master = pd.read_excel(MASTER_FILE_PATH, engine='openpyxl')
                else:
                    df_master = pd.DataFrame(columns=['Numele echipei', 'Scor Final'])

                current_scores = st.session_state.teams[['Echipă', 'Total']].copy()
                current_scores = current_scores.rename(columns={'Echipă': 'Numele echipei', 'Total': edition_col})

                # Avoid duplicate columns if pushing the same edition twice
                if edition_col in df_master.columns:
                    df_master = df_master.drop(columns=[edition_col])

                df_master = pd.merge(df_master, current_scores, on='Numele echipei', how='outer')

                # Re-sort columns: Name, then Editions in order, then Final Score
                cols = df_master.columns.tolist()
                edition_cols = [c for c in cols if "Ediția" in c]
                edition_cols.sort(key=lambda x: int(re.search(r'\d+', x).group()))
                df_master = df_master[['Numele echipei'] + edition_cols + ['Scor Final']]

                # Fill gaps and calculate total across all editions
                for col in edition_cols:
                    df_master[col] = df_master[col].fillna("-")

                def calculate_total(row):
                    total_val = 0
                    for col in edition_cols:
                        val = row[col]
                        if isinstance(val, (int, float)):
                            total_val += val
                    return total_val

                df_master['Scor Final'] = df_master.apply(calculate_total, axis=1)
                df_master = df_master.sort_values(by='Scor Final', ascending=False)

                try:
                    with pd.ExcelWriter(MASTER_FILE_PATH, engine='xlsxwriter') as writer:
                        df_master.to_excel(writer, index=False, sheet_name='Overall Ranking')
                    st.toast(f"✅ Master Sheet Updated", icon='🧛')
                except PermissionError:
                    st.error(f"❌ ACCES REFUZAT: Închide Excel-ul și apasă PUSH din nou!")

    num_cfg = st.column_config.NumberColumn(format="%.2f", step=0.01)
    edited_df = st.data_editor(
        st.session_state.teams,
        column_config={
            "Echipă": st.column_config.TextColumn("Echipă", width="medium"),
            "Joker": st.column_config.SelectboxColumn("Joker", options=["Niciuna", "R1", "R2", "R3", "R4", "R5"]),
            "Pariu (±)": num_cfg,
            "Total": st.column_config.NumberColumn("Total", disabled=True, format="%.2f"),
            "R1": num_cfg, "R2": num_cfg, "R3": num_cfg, "R4": num_cfg, "R5": num_cfg,
        },
        hide_index=False, use_container_width=True, num_rows="dynamic"
    )
    if not edited_df.equals(st.session_state.teams):
        st.session_state.teams = run_calculation(edited_df)
        st.rerun()


# --- 2. FULLSCREEN REVEAL ENGINE ---
def show_leaderboard_reveal():
    st.header("📽️ Reveal Clasament")
    if st.session_state.teams.empty:
        st.warning("Adaugă echipe în Scoring Matrix pentru a începe.")
        return

    leaderboard = st.session_state.teams[['Echipă', 'Total']].sort_values(by='Total', ascending=True).reset_index(
        drop=True)
    total_teams = len(leaderboard)

    if 'reveal_step' not in st.session_state: st.session_state.reveal_step = 0

    col1, col2 = st.columns([1, 5])
    if col1.button("⏪ Reset"):
        st.session_state.reveal_step = 0
        st.rerun()

    step = st.session_state.reveal_step
    if step >= total_teams:
        st.success("Reveal Complet!")
        return

    current_team = leaderboard.iloc[step]
    rank = total_teams - step
    name_color = "#ffffff"
    if rank == 1:
        name_color = "#FFD700"
    elif rank == 2:
        name_color = "#C0C0C0"
    elif rank == 3:
        name_color = "#CD7F32"

    st.markdown(f"""
        <style>
            .reveal-container {{
                background-color: #000000; padding: 50px 80px; border-radius: 20px; 
                border: 8px solid #00004a; text-align: center; min-height: 55vh; 
                display: flex; flex-direction: column; justify-content: center; align-items: center; 
                box-shadow: 0 0 40px #000080;
            }}
        </style>
        <div class="reveal-container">
            <h2 style="color: #666666; font-size: 30px; margin: 0; font-family: 'Chromium One', sans-serif;">LOCUL</h2>
            <h1 style="color: #ffffff; font-size: 110px; margin: 0; line-height: 1; font-family: 'Chromium One', sans-serif;">{rank}</h1>
            <div style="width: 250px; height: 3px; background-color: #000080; margin: 20px 0;"></div>
            <h2 style="color: {name_color}; font-size: 80px; font-weight: bold; margin: 5px 0;">{current_team['Echipă']}</h2>
            <h3 style="color: #ffffff; font-size: 50px; margin: 0;">{current_team['Total']:.2f}</h3>
        </div>
    """, unsafe_allow_html=True)

    if st.button("Următorul Loc", use_container_width=True):
        st.session_state.reveal_step += 1
        st.rerun()


# --- 3. STYLE & IMAGE HELPERS ---
def place_smart_scaled_image(slide, img_path, target_center_x, target_center_y, max_w=Inches(5.5), max_h=Inches(4.5)):
    exts = ['.jpg', '.jpeg', '.png', '.avif', '.webp']
    final_path = img_path if img_path and os.path.exists(img_path) else None
    if not final_path and img_path:
        base = os.path.splitext(img_path)[0]
        for ext in exts:
            if os.path.exists(base + ext): final_path = base + ext; break
    if not final_path: return

    PPTX_SUPPORTED = {'JPEG', 'PNG', 'BMP', 'GIF', 'TIFF', 'WMF'}
    try:
        img = Image.open(final_path)
        actual_format = img.format
    except Exception:
        img, actual_format = None, None

    if img is not None and actual_format not in PPTX_SUPPORTED:
        img_io = io.BytesIO()
        img.save(img_io, format='PNG')
        img_io.seek(0)
        img_to_add = img_io
    else:
        img_to_add = final_path

    try:
        pic = slide.shapes.add_picture(img_to_add, 0, 0)
    except ValueError as e:
        raise ValueError(f"Unsupported image format for file '{final_path}': {e}") from e
    ratio = min(max_w / pic.width, max_h / pic.height)
    new_w, new_h = int(pic.width * ratio), int(pic.height * ratio)
    pic.width, pic.height, pic.left, pic.top = new_w, new_h, int(target_center_x - (new_w / 2)), int(
        target_center_y - (new_h / 2))


def add_styled_text(slide, text, left, top, width, height, font_size=None, is_qa=False, bold=False,
                    alignment=PP_ALIGN.CENTER, font_name=None, color_rgb=(255, 255, 255), force_single_line=True):
    if not text or str(text).strip() == "": return
    if font_size is None: font_size = get_dynamic_font_size(text, is_qa)
    txBox = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    tf = txBox.text_frame
    tf.word_wrap, tf.vertical_anchor = not force_single_line, MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text, p.alignment = str(text), alignment
    if len(p.runs) > 0:
        run = p.runs[0]
        if force_single_line:
            ratio = 1.4 if font_name == "Chromium One" else 1.9
            cur_size = font_size
            while (len(str(text)) * (cur_size / ratio)) > (width / 914400 * 72):
                cur_size -= 5
                if cur_size < 18: break
            font_size = cur_size
        run.font.size, run.font.bold = Pt(font_size), bold
        if font_name: run.font.name = font_name
        run.font.color.rgb = RGBColor(*color_rgb)


# --- 4. CORE GENERATOR ---
def generate_trivia_slides(df):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)

    def add_bg(slide, path=None):
        bg = path if path else os.path.join("images", "background.jpeg")
        if os.path.exists(bg): slide.shapes.add_picture(bg, 0, 0, width=prs.slide_width, height=prs.slide_height)

    def add_styled_slide(t="", s=""):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide)
        if t: add_styled_text(slide, t, Inches(0.5), Inches(0.5), prs.slide_width - Inches(1), Inches(3), font_size=150,
                              font_name="Chromium One", color_rgb=(255, 215, 0))
        if s: add_styled_text(slide, s, Inches(0.5), Inches(4), prs.slide_width - Inches(1), Inches(3), font_size=125,
                              font_name="Gladiola", color_rgb=(255, 255, 255))
        return slide

    for img in ["First Slide.jpeg", "Second Slide.jpeg", "Third Slide.jpeg", "Fourth Slide.jpeg"]:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide, os.path.join("images", "Slides", "Generic Slides", img))

    for r in sorted([r for r in df['Round'].unique() if r <= 5]):
        r_data = df[df['Round'] == r]
        base_name = r_data['Round_Name'].iloc[0]
        r_img = os.path.join("images", "Slides", "Round Slides", f"Round {r}.jpeg")
        if os.path.exists(r_img):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_bg(slide, r_img)
        else:
            add_styled_slide(f"Runda {r}", base_name)

        # --- QUESTION SLIDES LOOP ---
        for i, (_, row) in enumerate(r_data.iterrows(), 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            img_n = str(row.get('Picture')) if pd.notna(row.get('Picture')) else None
            if img_n and img_n.strip().lower() in ['nan', 'none', '']:
                img_n = None
            path = os.path.join("images", f"Round {r}", img_n) if img_n else None
            has_pic = False
            if path:
                if os.path.exists(path):
                    has_pic = True
                else:
                    base = os.path.splitext(path)[0]
                    for ext in ['.jpg', '.jpeg', '.png', '.avif', '.webp']:
                        if os.path.exists(base + ext):
                            has_pic = True
                            path = base + ext
                            break

            r3_q_img = None
            if r == 2:
                candidates_r2 = [
                    os.path.join("images", "Slides", "Round 2", "Question", f"Question {i}.jpeg"),
                    os.path.join("images", "Slides", "Round 2", "Question", f"Question {i}.jpg"),
                    os.path.join("images", "Slides", "Round 2", "Question.jpeg")
                ]
                q_bg = next((c for c in candidates_r2 if os.path.exists(c)), None)
                add_bg(slide, q_bg)
            elif r == 3:
                for candidate in [
                    os.path.join("images", "Slides", "Round 3", "Questions", f"{i}.jpeg"),
                    os.path.join("images", "Slides", "Round 3", f"{i}.jpeg")
                ]:
                    if os.path.exists(candidate):
                        r3_q_img = candidate
                        break
                add_bg(slide, r3_q_img if r3_q_img else None)
            elif r in [1, 4, 5]:
                r145_dirs = [
                    os.path.join("images", "Slides", "Round 1, 4, 5"),
                    os.path.join("images", "Slides", "Round 1,4,5")
                ]
                r145_dir = next((d for d in r145_dirs if os.path.exists(d)), r145_dirs[0])
                if has_pic:
                    candidates_r145 = [
                        os.path.join(r145_dir, "Question and Picture", f"Question and Picture {i}.jpeg"),
                        os.path.join(r145_dir, "Question and Picture", f"Question and Picture {i}.jpg"),
                        os.path.join(r145_dir, "Question and Picture", f"{i}.jpeg"),
                        os.path.join(r145_dir, "Question and Picture", f"{i}.jpg"),
                        os.path.join(r145_dir, "Question and Picture.jpeg")
                    ]
                else:
                    candidates_r145 = [
                        os.path.join(r145_dir, "Question", f"Question {i}.jpeg"),
                        os.path.join(r145_dir, "Question", f"Question {i}.jpg"),
                        os.path.join(r145_dir, "Question", f"{i}.jpeg"),
                        os.path.join(r145_dir, "Question", f"{i}.jpg"),
                        os.path.join(r145_dir, "Question.jpeg")
                    ]
                r145_bg = next((c for c in candidates_r145 if os.path.exists(c)), None)
                add_bg(slide, r145_bg)
            else:
                add_bg(slide)

            # AUDIO LOGIC (Always runs for Round 3)
            if r == 3:
                a_base = str(row['Question'])
                q_audio_folder = os.path.join("audio", "questions")
                a_exts = ['', '.mp3', '.wav', '.mpeg', '.m4a', '.mp4']
                found_q = next((os.path.join(q_audio_folder, os.path.splitext(a_base)[0] + e) for e in a_exts if
                                os.path.exists(os.path.join(q_audio_folder, os.path.splitext(a_base)[0] + e))), None)
                if found_q:
                    final_q_audio = found_q
                    if not found_q.lower().endswith('.mp3'):
                        try:
                            conv_q_path = os.path.splitext(found_q)[0] + "_converted.mp3"
                            if not os.path.exists(conv_q_path): AudioSegment.from_file(found_q).export(conv_q_path,
                                                                                                       format="mp3")
                            final_q_audio = conv_q_path
                        except Exception:
                            pass
                    slide.shapes.add_movie(os.path.abspath(final_q_audio), Inches(-1), Inches(-1), Inches(0.5),
                                           Inches(0.5)).name = "TriviaAudio"

            # QUESTION SLIDE LAYOUT LOGIC
            if r == 2:
                # 3 Images Side-by-Side (or 2 Images if C is not present)
                img_q_val = str(row.get('Question'))
                if not img_q_val or img_q_val.strip().lower() in ['poza', 'nan', 'none', '']:
                    img_q_val = f"{i}A"
                img_path_a = os.path.join("images", f"Round {r}", img_q_val)
                img_path_b = path if path else os.path.join("images", f"Round {r}", f"{i}B")
                img_path_c = os.path.join("images", f"Round {r}", f"{i}C")

                c_exists = any(os.path.exists(os.path.splitext(img_path_c)[0] + ext) for ext in ['', '.jpeg', '.jpg', '.png', '.avif', '.webp'])

                if c_exists:
                    place_smart_scaled_image(slide, img_path_a, Inches(2.65), Inches(3.7), max_w=Inches(3.3), max_h=Inches(3.6))
                    place_smart_scaled_image(slide, img_path_b, Inches(6.67), Inches(3.7), max_w=Inches(3.3), max_h=Inches(3.6))
                    place_smart_scaled_image(slide, img_path_c, Inches(10.68), Inches(3.7), max_w=Inches(3.3), max_h=Inches(3.6))
                else:
                    if img_path_a:
                        place_smart_scaled_image(slide, img_path_a, prs.slide_width * 0.25,
                                                 prs.slide_height / 2 + Inches(0.4), max_w=Inches(6.0), max_h=Inches(5.0))
                    if img_path_b:
                        place_smart_scaled_image(slide, img_path_b, prs.slide_width * 0.75,
                                                 prs.slide_height / 2 + Inches(0.4), max_w=Inches(6.0), max_h=Inches(5.0))

            elif r == 3:
                if not r3_q_img:
                    if "Continuă Versul" in base_name:
                        # Lyric on Left, Image on Right
                        ans = str(row.get('Answer', ''))
                        parts = ans.split("-", 1) if "-" in ans else [ans, ""]
                        lyric_q = parts[0].strip()

                        add_styled_text(slide, lyric_q, Inches(0.5), Inches(1.5), prs.slide_width / 2 - Inches(0.5), Inches(5),
                                        bold=True, is_qa=True, font_name="Gladiola", force_single_line=False)
                        if path:
                            place_smart_scaled_image(slide, path, prs.slide_width * 0.75, prs.slide_height / 2 + Inches(0.4),
                                                     max_w=Inches(5.5), max_h=Inches(5.0))
                    else:
                        if path:
                            place_smart_scaled_image(slide, path, prs.slide_width / 2, prs.slide_height / 2 + Inches(0.4),
                                                     max_w=Inches(5.0), max_h=Inches(5.0))

            else:
                # Standard Logic for R1, R4, R5 Question Slides
                if has_pic:
                    # Shape: Question and Picture
                    add_styled_text(slide, row['Question'], Pt(67.71), Pt(101.14), Pt(436.29), Pt(272.57),
                                    bold=True, is_qa=True, font_name="Gladiola", force_single_line=False)
                    place_smart_scaled_image(slide, path, Inches(10), prs.slide_height / 2 + Inches(0.5),
                                             max_w=Inches(5.0), max_h=Inches(4.5))
                else:
                    # Shape: Question
                    add_styled_text(slide, row['Question'], Pt(390), Pt(64.29), Pt(492), Pt(324),
                                    bold=True, is_qa=True, font_name="Gladiola", force_single_line=False)

        ans_img = os.path.join("images", "Slides", "Round Slides", f"Round {r} Answers.jpeg")
        if os.path.exists(ans_img):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_bg(slide, ans_img)
        else:
            add_styled_slide(f"Runda {r}", "Răspunsuri")

        # --- ANSWER SLIDES LOOP ---
        for i, (_, row) in enumerate(r_data.iterrows(), 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            img_n = str(row.get('Picture')) if pd.notna(row.get('Picture')) else None
            if img_n and img_n.strip().lower() in ['nan', 'none', '']:
                img_n = None
            path = os.path.join("images", f"Round {r}", img_n) if img_n else None
            has_pic = False
            if path:
                if os.path.exists(path):
                    has_pic = True
                else:
                    base = os.path.splitext(path)[0]
                    for ext in ['.jpg', '.jpeg', '.png', '.avif', '.webp']:
                        if os.path.exists(base + ext):
                            has_pic = True
                            path = base + ext
                            break

            if r == 2:
                candidates_r2 = [
                    os.path.join("images", "Slides", "Round 2", "Answer", f"Answer {i}.jpeg"),
                    os.path.join("images", "Slides", "Round 2", "Answer", f"Answer {i}.jpg"),
                    os.path.join("images", "Slides", "Round 2", "Answer.jpeg"),
                    os.path.join("images", "Slides", "Round 2", "Answers.jpeg")
                ]
                ans_bg = next((c for c in candidates_r2 if os.path.exists(c)), None)
                add_bg(slide, ans_bg)
            elif r == 3:
                r3_ans_bg = os.path.join("images", "Slides", "Round 3", "Answers", f"{i}.jpeg")
                if not os.path.exists(r3_ans_bg):
                    r3_ans_bg = os.path.join("images", "Slides", "Round 3", f"{i}.jpeg")
                add_bg(slide, r3_ans_bg if os.path.exists(r3_ans_bg) else None)
            elif r in [1, 4, 5]:
                r145_dirs = [
                    os.path.join("images", "Slides", "Round 1, 4, 5"),
                    os.path.join("images", "Slides", "Round 1,4,5")
                ]
                r145_dir = next((d for d in r145_dirs if os.path.exists(d)), r145_dirs[0])
                if has_pic:
                    candidates_r145 = [
                        os.path.join(r145_dir, "Answer and Picture", f"Answer and Picture {i}.jpeg"),
                        os.path.join(r145_dir, "Answer and Picture", f"Answer and Picture {i}.jpg"),
                        os.path.join(r145_dir, "Answer and Picture", f"{i}.jpeg"),
                        os.path.join(r145_dir, "Answer and Picture", f"{i}.jpg"),
                        os.path.join(r145_dir, "Answer and Picture.jpeg")
                    ]
                else:
                    candidates_r145 = [
                        os.path.join(r145_dir, "Answer", f"Answer {i}.jpeg"),
                        os.path.join(r145_dir, "Answer", f"Answer {i}.jpg"),
                        os.path.join(r145_dir, "Answer", f"{i}.jpeg"),
                        os.path.join(r145_dir, "Answer", f"{i}.jpg"),
                        os.path.join(r145_dir, "Answer.jpeg")
                    ]
                r145_bg = next((c for c in candidates_r145 if os.path.exists(c)), None)
                add_bg(slide, r145_bg)
            else:
                add_bg(slide)

            # ANSWER SLIDE LAYOUT LOGIC
            # --- ADAPTIVE ROUND 3 ANSWER LAYOUT ---
            if r == 3:
                # Lyric Continuation on Left, Image on Right
                if "Continuă Versul" in base_name:
                    ans = str(row.get('Answer', ''))
                    parts = ans.split("-", 1) if "-" in ans else ["", ans]
                    lyric_a = parts[1].strip() if len(parts) > 1 else ans.strip()
                    add_styled_text(slide, lyric_a, Inches(0.5), Inches(1.5), prs.slide_width / 2 - Inches(0.5),
                                    Inches(5), bold=True, is_qa=True, font_name="Gladiola", force_single_line=False)
                    if path:
                        place_smart_scaled_image(slide, path, prs.slide_width * 0.75,
                                                 prs.slide_height / 2 + Inches(0.4), max_w=Inches(5.5),
                                                 max_h=Inches(5.0))

                # Classic: 3-Zone Layout (Song - Image - Artist)
                else:
                    cw, vt = prs.slide_width / 3, (prs.slide_height - Inches(2.5)) / 2
                    ans = str(row['Answer'])
                    parts = ans.split("-", 1) if "-" in ans else [ans, ""]
                    # Left Zone: Song Name
                    add_styled_text(slide, parts[0].strip(), Inches(0.2), vt, cw - Inches(0.4), Inches(2.5),
                                    font_size=48, bold=True, font_name="Gladiola", force_single_line=False)
                    # Middle Zone: Image
                    if path:
                        place_smart_scaled_image(slide, path, prs.slide_width / 2, prs.slide_height / 2,
                                                 max_w=cw, max_h=Inches(5))
                    # Right Zone: Artist Name
                    if len(parts) > 1:
                        add_styled_text(slide, parts[1].strip(), prs.slide_width - cw + Inches(0.2), vt,
                                        cw - Inches(0.4), Inches(2.5), font_size=48, bold=True, font_name="Gladiola",
                                        force_single_line=False)

                # --- AUDIO ENGINE (SHARED BY BOTH MODES) ---
                a_base = str(row['Question'])
                ans_audio_folder = os.path.join("audio", "answers")
                a_exts = ['', '.mp3', '.wav', '.mpeg', '.m4a', '.mp4']
                found_ans = next((os.path.join(ans_audio_folder, os.path.splitext(a_base)[0] + e) for e in a_exts if
                                  os.path.exists(os.path.join(ans_audio_folder, os.path.splitext(a_base)[0] + e))),
                                 None)
                if found_ans:
                    final_ans_audio = found_ans
                    if not found_ans.lower().endswith('.mp3'):
                        try:
                            conv_ans_path = os.path.splitext(found_ans)[0] + "_ans_converted.mp3"
                            if not os.path.exists(conv_ans_path):
                                AudioSegment.from_file(found_ans).export(conv_ans_path, format="mp3")
                            final_ans_audio = conv_ans_path
                        except Exception:
                            pass
                    slide.shapes.add_movie(os.path.abspath(final_ans_audio), Inches(-1), Inches(-1), Inches(0.5),
                                           Inches(0.5)).name = "TriviaAudio"
            elif r == 2:
                # 3 Images Side-by-Side + Text at Bottom (or 2 Images if C is not present)
                img_q_val = str(row.get('Question'))
                if not img_q_val or img_q_val.strip().lower() in ['poza', 'nan', 'none', '']:
                    img_q_val = f"{i}A"
                img_path_a = os.path.join("images", f"Round {r}", img_q_val)
                img_path_b = path if path else os.path.join("images", f"Round {r}", f"{i}B")
                img_path_c = os.path.join("images", f"Round {r}", f"{i}C")

                c_exists = any(os.path.exists(os.path.splitext(img_path_c)[0] + ext) for ext in ['', '.jpeg', '.jpg', '.png', '.avif', '.webp'])

                if c_exists:
                    place_smart_scaled_image(slide, img_path_a, Inches(2.65), Inches(3.7), max_w=Inches(3.3), max_h=Inches(3.6))
                    place_smart_scaled_image(slide, img_path_b, Inches(6.67), Inches(3.7), max_w=Inches(3.3), max_h=Inches(3.6))
                    place_smart_scaled_image(slide, img_path_c, Inches(10.68), Inches(3.7), max_w=Inches(3.3), max_h=Inches(3.6))
                else:
                    if img_path_a:
                        place_smart_scaled_image(slide, img_path_a, prs.slide_width * 0.25,
                                                 prs.slide_height / 2 - Inches(0.5), max_w=Inches(6.0), max_h=Inches(4.5))
                    if img_path_b:
                        place_smart_scaled_image(slide, img_path_b, prs.slide_width * 0.75,
                                                 prs.slide_height / 2 - Inches(0.5), max_w=Inches(6.0), max_h=Inches(4.5))

                add_styled_text(slide, f"Răspuns: {row['Answer']}", 0, Inches(6.2), prs.slide_width, Inches(1.0),
                                font_size=54, bold=True, font_name="Gladiola", force_single_line=False)

            else:
                # Standard Logic for R1, R4, R5 Answer Slides
                if has_pic:
                    # Shape: Answer and Picture Question & Answer and Picture Answer
                    ans_text = f"Răspuns: {row['Answer']}"
                    ans_font_size = 32 if len(ans_text) > 40 else (38 if len(ans_text) > 28 else 44)
                    add_styled_text(slide, row['Question'], Pt(72.86), Pt(76.29), Pt(431.14), Pt(204.51),
                                    is_qa=True, font_name="Gladiola", force_single_line=False)
                    add_styled_text(slide, ans_text, Pt(72.86), Pt(330), Pt(418.89), Pt(73.71),
                                    font_size=ans_font_size, bold=True, is_qa=True, font_name="Gladiola", force_single_line=False)
                    place_smart_scaled_image(slide, path, Inches(10), prs.slide_height / 2 + Inches(0.5),
                                             max_w=Inches(5.0), max_h=Inches(4.5))
                else:
                    # Shape: Answer Question & Answer Answer
                    ans_text = f"Răspuns: {row['Answer']}"
                    ans_font_size = 32 if len(ans_text) > 40 else (38 if len(ans_text) > 28 else 44)
                    add_styled_text(slide, row['Question'], Pt(396), Pt(69.43), Pt(483.43), Pt(200.57),
                                    is_qa=True, font_name="Gladiola", force_single_line=False)
                    add_styled_text(slide, ans_text, Pt(402.86), Pt(330.86), Pt(468), Pt(93.94),
                                    font_size=ans_font_size, bold=True, is_qa=True, font_name="Gladiola", force_single_line=False)

        if r in [3, 5]:
            break_filename = "15 Min Break Slide.jpeg" if r == 3 else "10 Min Break Slide.jpeg"
            break_img = os.path.join("images", "Slides", "Break Slides", break_filename)
            lb_img = os.path.join("images", "Slides", "Break Slides", "Leaderboard Slide.jpeg")
            if os.path.exists(break_img):
                b_slide = prs.slides.add_slide(prs.slide_layouts[6])
                add_bg(b_slide, break_img)
            else:
                add_styled_slide("Pauză", "Rezultatele momentului" if r == 3 else "Situația înainte de Pariu")

            if os.path.exists(lb_img):
                lb_slide = prs.slides.add_slide(prs.slide_layouts[6])
                add_bg(lb_slide, lb_img)

    # --- WAGER ---
    p_data = df[df['Round'] == 6]
    if not p_data.empty:
        p_row = p_data.iloc[0]
        wager_img = os.path.join("images", "Slides", "Round Slides", "Wager.jpeg")
        if os.path.exists(wager_img):
            w_slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_bg(w_slide, wager_img)
        else:
            add_styled_slide("Pariul", "Miza crește...")

        # Question slide
        sq = prs.slides.add_slide(prs.slide_layouts[6])
        w_q_bg = os.path.join("images", "Slides", "Wager", "Question.jpeg")
        add_bg(sq, w_q_bg if os.path.exists(w_q_bg) else None)

        q_text = str(p_row.get('Question', '')).strip()
        pattern = r'^(.*?)(?:\r?\n|\s+)(?:A[\).:]\s*)(.*?)(?:\r?\n|\s+)(?:B[\).:]\s*)(.*?)(?:\r?\n|\s+)(?:C[\).:]\s*)(.*?)(?:\r?\n|\s+)(?:D[\).:]\s*)(.*?)$'
        m_opts = re.search(pattern, q_text, re.DOTALL)

        if os.path.exists(w_q_bg):
            if m_opts:
                stem, opt_a, opt_b, opt_c, opt_d = m_opts.groups()
                add_styled_text(sq, stem.strip(), Inches(1.3), Inches(2.0), Inches(10.7), Inches(2.1), font_size=34, bold=True,
                                is_qa=True, font_name="Gladiola", force_single_line=False)
                add_styled_text(sq, opt_a.strip(), Inches(1.8), Inches(4.9), Inches(4.5), Inches(0.8), font_size=32, bold=True,
                                alignment=PP_ALIGN.LEFT, font_name="Gladiola", force_single_line=False)
                add_styled_text(sq, opt_b.strip(), Inches(8.0), Inches(4.9), Inches(4.5), Inches(0.8), font_size=32, bold=True,
                                alignment=PP_ALIGN.LEFT, font_name="Gladiola", force_single_line=False)
                add_styled_text(sq, opt_c.strip(), Inches(1.8), Inches(6.3), Inches(4.5), Inches(0.8), font_size=32, bold=True,
                                alignment=PP_ALIGN.LEFT, font_name="Gladiola", force_single_line=False)
                add_styled_text(sq, opt_d.strip(), Inches(8.0), Inches(6.3), Inches(4.5), Inches(0.8), font_size=32, bold=True,
                                alignment=PP_ALIGN.LEFT, font_name="Gladiola", force_single_line=False)
            else:
                add_styled_text(sq, q_text, Inches(1.3), Inches(2.0), Inches(10.7), Inches(4.5), font_size=36, bold=True,
                                is_qa=True, font_name="Gladiola", force_single_line=False)
        else:
            add_styled_text(sq, "Pariul: Întrebarea", Inches(0.5), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                            alignment=PP_ALIGN.LEFT, font_name="Gladiola", color_rgb=(255, 215, 0))
            w_img = glob.glob(os.path.join("images", "Wager", "*.*"))
            if w_img:
                add_styled_text(sq, p_row['Question'], Inches(0.5), Inches(1.0), Inches(6.0), Inches(6.0), bold=True,
                                is_qa=True, font_name="Gladiola", force_single_line=False)
                place_smart_scaled_image(sq, w_img[0], Inches(10), prs.slide_height / 2 + Inches(0.5), max_w=Inches(5.0),
                                         max_h=Inches(4.5))
            else:
                add_styled_text(sq, p_row['Question'], Inches(1.0), Inches(1.5), Inches(11.3), Inches(5.0), bold=True,
                                is_qa=True, font_name="Gladiola", force_single_line=False)

        # Answer slide
        ans_raw = str(p_row.get('Answer', '')).strip()
        m_ans = re.search(r'\b([A-D])\b', ans_raw, re.IGNORECASE)
        if not m_ans:
            m_ans = re.search(r'([A-D])', ans_raw, re.IGNORECASE)
        ans_letter = m_ans.group(1).upper() if m_ans else 'A'

        w_a_bg = os.path.join("images", "Slides", "Wager", f"Answer {ans_letter}.jpeg")
        sa = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(sa, w_a_bg if os.path.exists(w_a_bg) else None)

        if os.path.exists(w_a_bg):
            stem_text = stem.strip() if m_opts else q_text
            add_styled_text(sa, stem_text, Inches(1.3), Inches(2.35), Inches(10.7), Inches(2.1), font_size=34, bold=True,
                            is_qa=True, font_name="Gladiola", force_single_line=False)
            clean_ans = re.sub(rf'^{ans_letter}[\).:\s-]+', '', ans_raw, flags=re.IGNORECASE).strip()
            if not clean_ans:
                clean_ans = ans_raw
            ans_font_size = 22 if len(clean_ans) > 70 else (28 if len(clean_ans) > 40 else 34)
            add_styled_text(sa, clean_ans, Inches(3.4), Inches(5.5), Inches(7.4), Inches(1.0), font_size=ans_font_size,
                            bold=True, alignment=PP_ALIGN.LEFT, font_name="Gladiola", force_single_line=False)
        else:
            add_styled_text(sa, "Pariul: Răspuns", Inches(0.5), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                            alignment=PP_ALIGN.LEFT, font_name="Gladiola", color_rgb=(255, 215, 0))
            add_styled_text(sa, f"Răspuns: {p_row['Answer']}", Inches(1.0), Inches(1.5), Inches(11.3), Inches(5.0),
                            font_size=54, bold=True, is_qa=True, font_name="Gladiola", force_single_line=False)

        last_img = os.path.join("images", "Slides", "Generic Slides", "Last Slide.jpeg")
        if os.path.exists(last_img):
            last_slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_bg(last_slide, last_img)
        else:
            add_styled_slide("Rezultate Finale", "Câștigătorii Transylvania Trivia")

    prs.save("Transylvania_Trivia.pptx")
    return "Transylvania_Trivia.pptx"


# --- 5. AGENT STUDIO (HUMAN-IN-THE-LOOP AUTOMATION) ---
def render_agent_studio_tab():
    st.header("🤖 Agent Studio - HITL Trivia Orchestrator")
    st.markdown("Generați și curatoriați automat întrebări pentru Transylvania Trivia folosind **google-antigravity AI**.")

    import agent_orchestrator
    saved_key = agent_orchestrator.get_saved_api_key()

    with st.expander("⚙️ Configurare Generare AI", expanded=True):
        col_t4, col_t5 = st.columns(2)
        with col_t4:
            theme_r4 = st.text_input("Tema Runda 4", value="Istoria României")
        with col_t5:
            theme_r5 = st.text_input("Tema Runda 5", value="Film și Cinema")

        st.caption("🎵 **Runda 3 (Muzică):** Întrebările și melodiile sunt preluate automat din `questions.csv`, iar imaginile sunt căutate și descărcate automat.")

        if st.button("🚀 Generare Candidați (15 întrebări / rundă)", use_container_width=True, type="primary"):
            if not saved_key:
                st.error("❌ Nu a fost găsită nicio cheie API Gemini! Asigurați-vă că ați setat GEMINI_API_KEY în fișierul .env sau în agent_orchestrator.py.")
            else:
                round_configs = [
                    {"round_nr": 1, "round_name": "Cultură Generală", "count": 15},
                    {"round_nr": 2, "round_name": "Ce/Cine se află în această poză?", "count": 15},
                    {"round_nr": 4, "round_name": theme_r4 if theme_r4 else "Runda Tematică 1", "count": 15},
                    {"round_nr": 5, "round_name": theme_r5 if theme_r5 else "Runda Tematică 2", "count": 15},
                    {"round_nr": 6, "round_name": "Pariul", "count": 5},
                ]

                progress_bar = st.progress(0.0)
                status_text = st.empty()

                def p_callback(cur, total, msg):
                    progress_bar.progress(cur / total)
                    status_text.info(f"⏳ {msg}")

                try:
                    import agent_orchestrator
                    status_text.info("🤖 Se generează întrebările cu google.antigravity Agent...")
                    results = asyncio.run(agent_orchestrator.run_full_generation_pipeline(
                        api_key=saved_key,
                        round_configs=round_configs,
                        progress_callback=p_callback,
                        model_name=agent_orchestrator.DEFAULT_MODEL
                    ))


                    
                    status_text.info("🖼️ Se descarcă și se stochează imaginile în cache...")
                    agent_orchestrator.download_candidate_images(results, progress_callback=p_callback)
                    
                    st.session_state.generated_candidates = results
                    
                    # Initialize selection state (default top 10 for R1-R5, top 1 for R6)
                    sel_state = {}
                    for r_nr, cands in results.items():
                        target_sel = 1 if r_nr == 6 else 10
                        sel_state[r_nr] = [True if i < target_sel else False for i in range(len(cands))]
                    st.session_state.selected_candidate_indices = sel_state

                    progress_bar.progress(1.0)
                    status_text.success("✅ Generare finalizată cu succes! Revedeți candidații mai jos.")
                    st.rerun()
                except Exception as e:
                    status_text.error(f"❌ Eroare la generare: {e}")

    # --- CANDIDATE REVIEW & SELECTION UI ---
    if "generated_candidates" in st.session_state and st.session_state.generated_candidates:
        cands_by_round = st.session_state.generated_candidates
        sel_indices = st.session_state.get("selected_candidate_indices", {})

        st.markdown("---")
        st.subheader("🔍 Review & Curatare Candidați (HITL)")
        st.caption("Bifați întrebările dorite (10 per rundă / 1 pentru Pariu). Folosiți butonul 'Regenerare' sub imagini dacă selecția automată nu este optimă.")

        active_rounds = [1, 2, 3, 4, 5, 6]
        tab_titles = [
            f"Runda 1: Cultură Generală ({sum(sel_indices.get(1, []))}/10)",
            f"Runda 2: Poză ({sum(sel_indices.get(2, []))}/10)",
            f"Runda 3: Melodie ({sum(sel_indices.get(3, []))}/10)",
            f"Runda 4 ({sum(sel_indices.get(4, []))}/10)",
            f"Runda 5 ({sum(sel_indices.get(5, []))}/10)",
            f"Wager Pariu ({sum(sel_indices.get(6, []))}/1)"
        ]

        r_tabs = st.tabs(tab_titles)
        import agent_orchestrator

        for tab_idx, r_nr in enumerate(active_rounds):
            with r_tabs[tab_idx]:
                cands = cands_by_round.get(r_nr, [])
                r_sel = sel_indices.get(r_nr, [False] * len(cands))
                
                target_count = 1 if r_nr == 6 else 10
                curr_count = sum(r_sel)
                
                if curr_count == target_count:
                    st.success(f"✅ {curr_count}/{target_count} selectate pentru Runda {r_nr}")
                else:
                    st.warning(f"⚠️ {curr_count}/{target_count} selectate (trebuie exact {target_count})")

                if r_nr == 3:
                    st.info("🎵 **Runda 3 (Muzică):** Piesele și întrebările sunt preluate din `questions.csv`. Imaginile de mai jos au fost căutate și descărcate automat după artist și titlu. Puteți edita textul din 'Căutare imagine' și apăsa '🔄 Regenerare Imagine' dacă doriți o altă poză.")

                for c_idx, cand in enumerate(cands):
                    box_key = f"chk_{r_nr}_{c_idx}"
                    
                    is_checked = st.checkbox(
                        f"**Candidați #{c_idx+1}: {cand.get('answer')}**",
                        value=r_sel[c_idx] if c_idx < len(r_sel) else False,
                        key=box_key
                    )
                    r_sel[c_idx] = is_checked
                    
                    col_info, col_img = st.columns([1, 1] if r_nr == 2 else [3, 2])
                    with col_info:
                        if r_nr == 2:
                            st.write(f"**Întrebare:** `Poza`")
                        else:
                            st.write(f"**Întrebare:** {cand.get('question')}")
                        st.write(f"**Răspuns:** {cand.get('answer')}")
                        if r_nr == 2:
                            c_qa, c_qb, c_qc = st.columns(3)
                            with c_qa:
                                q_a_val = st.text_input(
                                    "Căutare imagine A:",
                                    value=cand.get("image_query_a", cand.get("answer", "")),
                                    key=f"txt_img_{r_nr}_{c_idx}_A"
                                )
                                cand["image_query_a"] = q_a_val
                            with c_qb:
                                q_b_val = st.text_input(
                                    "Căutare imagine B:",
                                    value=cand.get("image_query_b", cand.get("answer", "")),
                                    key=f"txt_img_{r_nr}_{c_idx}_B"
                                )
                                cand["image_query_b"] = q_b_val
                            with c_qc:
                                q_c_val = st.text_input(
                                    "Căutare imagine C:",
                                    value=cand.get("image_query_c", cand.get("answer", "")),
                                    key=f"txt_img_{r_nr}_{c_idx}_C"
                                )
                                cand["image_query_c"] = q_c_val
                        else:
                            q_val = st.text_input(
                                "Căutare imagine:",
                                value=cand.get("image_query", cand.get("answer", "")),
                                key=f"txt_img_{r_nr}_{c_idx}"
                            )
                            cand["image_query"] = q_val

                    with col_img:
                        if r_nr == 2:
                            col_a, col_b, col_c = st.columns(3)
                            with col_a:
                                pa = cand.get("cached_img_a")
                                if pa and os.path.exists(pa):
                                    st.image(pa, caption="Poză A (Stânga)", use_container_width=True)
                                else:
                                    st.caption("Imagine lipsă A")
                                if st.button("🔄 Regenerare A", key=f"regen_{r_nr}_{c_idx}_A"):
                                    if agent_orchestrator.regenerate_candidate_image(cand, r_nr, "A", custom_query=cand.get("image_query_a")):
                                        st.toast(f"Imagine A schimbată pentru #{c_idx+1}", icon="🖼️")
                                        st.rerun()
                                    else:
                                        st.warning("Nu s-a putut descărca altă imagine.")
                            with col_b:
                                pb = cand.get("cached_img_b")
                                if pb and os.path.exists(pb):
                                    st.image(pb, caption="Poză B (Mijloc)", use_container_width=True)
                                else:
                                    st.caption("Imagine lipsă B")
                                if st.button("🔄 Regenerare B", key=f"regen_{r_nr}_{c_idx}_B"):
                                    if agent_orchestrator.regenerate_candidate_image(cand, r_nr, "B", custom_query=cand.get("image_query_b")):
                                        st.toast(f"Imagine B schimbată pentru #{c_idx+1}", icon="🖼️")
                                        st.rerun()
                                    else:
                                        st.warning("Nu s-a putut descărca altă imagine.")
                            with col_c:
                                pc = cand.get("cached_img_c")
                                if pc and os.path.exists(pc):
                                    st.image(pc, caption="Poză C (Dreapta)", use_container_width=True)
                                else:
                                    st.caption("Imagine lipsă C")
                                if st.button("🔄 Regenerare C", key=f"regen_{r_nr}_{c_idx}_C"):
                                    if agent_orchestrator.regenerate_candidate_image(cand, r_nr, "C", custom_query=cand.get("image_query_c")):
                                        st.toast(f"Imagine C schimbată pentru #{c_idx+1}", icon="🖼️")
                                        st.rerun()
                                    else:
                                        st.warning("Nu s-a putut descărca altă imagine.")
                        else:
                            p_img = cand.get("cached_img")
                            if p_img and os.path.exists(p_img):
                                st.image(p_img, caption="Imagine Căutată", width=220)
                            else:
                                st.caption("Imagine lipsă")
                            if st.button("🔄 Regenerare Imagine", key=f"regen_{r_nr}_{c_idx}"):
                                if agent_orchestrator.regenerate_candidate_image(cand, r_nr, "main", custom_query=cand.get("image_query")):
                                    st.toast(f"Imagine schimbată pentru #{c_idx+1}", icon="🖼️")
                                    st.rerun()
                                else:
                                    st.warning("Nu s-a putut descărca altă imagine.")


                    st.markdown("---")

                sel_indices[r_nr] = r_sel

                if r_nr != 3:
                    btn_label = f"➕ Generează încă 10 întrebări pentru Runda {r_nr}" if r_nr != 6 else "➕ Generează încă 5 întrebări pentru Pariu"
                    if st.button(btn_label, key=f"btn_more_{r_nr}", use_container_width=True):
                        status_more = st.empty()
                        with st.spinner(f"Se generează întrebări noi pentru Runda {r_nr}..."):
                            count_to_gen = 5 if r_nr == 6 else 10
                            round_name = cands[0].get("round_name", f"Runda {r_nr}") if cands else f"Runda {r_nr}"
                            new_cands = asyncio.run(agent_orchestrator.generate_more_candidates_for_round(
                                api_key=saved_key,
                                round_nr=r_nr,
                                round_name=round_name,
                                count=count_to_gen,
                                existing_candidates=cands,
                                model_name=agent_orchestrator.DEFAULT_MODEL,
                                status_callback=lambda msg: status_more.info(msg)
                            ))

                            if new_cands:
                                agent_orchestrator.download_additional_candidate_images(r_nr, new_cands, start_index=len(cands))
                                cands.extend(new_cands)
                                r_sel.extend([False] * len(new_cands))
                                sel_indices[r_nr] = r_sel
                                st.session_state.generated_candidates[r_nr] = cands
                                st.session_state.selected_candidate_indices = sel_indices
                                st.toast(f"✅ Au fost adăugate {len(new_cands)} întrebări noi!", icon="🎉")
                                st.rerun()
                            else:
                                st.error("Nu s-au putut genera întrebări adiționale.")


        st.session_state.selected_candidate_indices = sel_indices

        # --- CONFIRMATION & EXECUTION ---
        st.markdown("### 🏆 Confirmare & Generare Finală")
        
        valid_counts = True
        validation_msgs = []
        for r_nr in [1, 2, 3, 4, 5, 6]:
            expected = 1 if r_nr == 6 else 10
            actual = sum(sel_indices.get(r_nr, []))
            if actual != expected:
                valid_counts = False
                validation_msgs.append(f"Runda {r_nr}: {actual}/{expected} selectate")

        if not valid_counts:
            st.warning("⚠️ Asigurați-vă că ați selectat exact numărul cerut de întrebări:\n" + ", ".join(validation_msgs))

        if st.button("Confirm", use_container_width=True, type="primary"):
            if not valid_counts:
                st.error("Nu puteți confirma până nu selectați exact 10 întrebări per rundă (și 1 pentru Pariu).")
            else:
                with st.spinner("Se salvează istoricul, se organizează imaginile și se construiește prezentarea PowerPoint..."):
                    selected_by_round = {}
                    unselected_by_round = {}

                    for r_nr, cands in cands_by_round.items():
                        r_sel = sel_indices.get(r_nr, [])
                        selected_by_round[r_nr] = [cands[i] for i, sel in enumerate(r_sel) if sel]
                        unselected_by_round[r_nr] = [cands[i] for i, sel in enumerate(r_sel) if not sel]

                    try:
                        pptx_path, total_qs = agent_orchestrator.confirm_and_finalize(
                            selected_by_round,
                            unselected_by_round,
                            generator_fn=generate_trivia_slides
                        )
                        st.balloons()
                        st.success(f"🎉 Succes! Am salvat {total_qs} întrebări în questions.csv și am generat '{pptx_path}'.")
                        with open(pptx_path, "rb") as f:
                            st.download_button("📥 Descarcă Prezentarea PowerPoint (.pptx)", f, file_name=pptx_path, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", use_container_width=True)
                    except Exception as e:
                        st.error(f"❌ Eroare la procesare: {e}")


# --- 4. APP EXECUTION ---
tab_cmd, tab_score, tab_reveal, tab_agent = st.tabs(["🎮 Trivia Command", "🏆 Scoring Matrix", "📽️ Leaderboard Presentation", "🤖 Agent Studio"])
with tab_cmd:
    st.title("🧛 Trivia Production")
    uploaded = st.file_uploader("Load CSV", type="csv", key="production_csv_uploader")
    if uploaded and st.button("Generate Slides"):
        path = generate_trivia_slides(pd.read_csv(uploaded))
        with open(path, "rb") as f: st.download_button("Download", f, file_name=path)
with tab_score: manage_scores()
with tab_reveal: show_leaderboard_reveal()
with tab_agent: render_agent_studio_tab()

