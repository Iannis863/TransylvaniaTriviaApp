import sys
import os

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
def manage_scores():
    st.header("🏆 Trivia Scoring Matrix")
    COLS = ['Echipă', 'R1', 'R2', 'R3', 'R4', 'R5', 'Joker', 'Pariu (±)', 'Total']
    if 'teams' not in st.session_state:
        st.session_state.teams = pd.DataFrame(columns=COLS)

    with st.expander("➕ Adaugă Echipă Nouă", expanded=True):
        t_name = st.text_input("Nume Echipă", key="new_team_name_input")
        if st.button("Înregistrează Echipa"):
            if t_name:
                new_t = pd.DataFrame(
                    [{'Echipă': t_name, 'R1': 0.00, 'R2': 0.00, 'R3': 0.00, 'R4': 0.00, 'R5': 0.00, 'Joker': "Niciuna",
                      'Pariu (±)': 0.00, 'Total': 0.00}])
                st.session_state.teams = pd.concat([st.session_state.teams, new_t], ignore_index=True)
                st.rerun()

    edited_df = st.data_editor(st.session_state.teams, use_container_width=True, num_rows="dynamic")
    if not edited_df.equals(st.session_state.teams):
        st.session_state.teams = edited_df
        st.rerun()


# --- 2. STYLE & IMAGE HELPERS ---
def place_smart_scaled_image(slide, img_path, target_center_x, target_center_y, max_w=Inches(5.5), max_h=Inches(4.5)):
    exts = ['.jpg', '.jpeg', '.png', '.avif', '.webp']
    final_path = img_path if img_path and os.path.exists(img_path) else None
    if not final_path and img_path:
        base = os.path.splitext(img_path)[0]
        for ext in exts:
            if os.path.exists(base + ext): final_path = base + ext; break
    if not final_path: return

    if os.path.splitext(final_path)[1].lower() in ['.webp', '.avif']:
        img = Image.open(final_path)
        img_io = io.BytesIO()
        img.save(img_io, format='PNG')
        img_io.seek(0)
        img_to_add = img_io
    else:
        img_to_add = final_path

    pic = slide.shapes.add_picture(img_to_add, 0, 0)
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


# --- 3. CORE GENERATOR ---
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
                              font_name="Chromium One", color_rgb=(0, 0, 255))
        if s: add_styled_text(slide, s, Inches(0.5), Inches(4), prs.slide_width - Inches(1), Inches(3), font_size=125,
                              font_name="Gladiola", color_rgb=(173, 216, 230))
        return slide

    for img in ["First Slide.png", "Second Slide.png", "Third Slide.png"]:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(slide, os.path.join("images", "Slides", img))

    for r in sorted([r for r in df['Round'].unique() if r <= 5]):
        r_data = df[df['Round'] == r]
        base_name = r_data['Round_Name'].iloc[0]
        add_styled_slide(f"Runda {r}", base_name)
        for i, (_, row) in enumerate(r_data.iterrows(), 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_bg(slide)
            add_styled_text(slide, row['Round_Name'] if r == 2 else base_name, Inches(0.5), Inches(0.2), Inches(5),
                            Inches(0.8), font_size=44, alignment=PP_ALIGN.LEFT, font_name="Gladiola",
                            color_rgb=(0, 0, 255))
            add_styled_text(slide, f"Întrebarea {i}", Inches(7.8), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                            alignment=PP_ALIGN.RIGHT, font_name="Gladiola", color_rgb=(0, 0, 255))

            img_n = str(row.get('Picture')) if pd.notna(row.get('Picture')) else None
            path = os.path.join("images", f"Round {r}", img_n) if img_n else None

            if r == 3:
                a_base = str(row['Question'])
                a_exts = ['', '.mp3', '.wav', '.mpeg', '.m4a', '.mp4']
                found = next((os.path.join("audio", os.path.splitext(a_base)[0] + e) for e in a_exts if
                              os.path.exists(os.path.join("audio", os.path.splitext(a_base)[0] + e))), None)
                if found:
                    final_audio = found
                    if not found.lower().endswith('.mp3'):
                        try:
                            conv_path = os.path.splitext(found)[0] + "_converted.mp3"
                            if not os.path.exists(conv_path):
                                st.toast(f"Converting {os.path.basename(found)}...")
                                AudioSegment.from_file(found).export(conv_path, format="mp3")
                            final_audio = conv_path
                        except Exception as e:
                            st.error(f"Audio Error: {e}")
                    slide.shapes.add_movie(os.path.abspath(final_audio), Inches(-1), Inches(-1), Inches(0.5),
                                           Inches(0.5)).name = "TriviaAudio"

            if r in [2, 3]:
                if path: place_smart_scaled_image(slide, path, prs.slide_width / 2, prs.slide_height / 2 + Inches(0.4),
                                                  max_w=Inches(9.0), max_h=Inches(5.0))
            else:
                add_styled_text(slide, row['Question'], Inches(0.5), Inches(1.5), Inches(6), Inches(5), bold=True,
                                is_qa=True, font_name="Gladiola", force_single_line=False)
                if path: place_smart_scaled_image(slide, path, Inches(10), prs.slide_height / 2 + Inches(0.5),
                                                  max_w=Inches(5.0), max_h=Inches(4.5))

        add_styled_slide(f"Runda {r}", "Răspunsuri")
        for i, (_, row) in enumerate(r_data.iterrows(), 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_bg(slide)
            add_styled_text(slide, row['Round_Name'] if r == 2 else base_name, Inches(0.5), Inches(0.2), Inches(5),
                            Inches(0.8), font_size=44, alignment=PP_ALIGN.LEFT, font_name="Gladiola",
                            color_rgb=(0, 0, 255))
            add_styled_text(slide, f"Răspuns {i}", Inches(7.8), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                            alignment=PP_ALIGN.RIGHT, font_name="Gladiola", color_rgb=(0, 0, 255))
            img_n = str(row.get('Picture')) if pd.notna(row.get('Picture')) else None
            path = os.path.join("images", f"Round {r}", img_n) if img_n else None

            if r == 3:
                cw, vt = prs.slide_width / 3, (prs.slide_height - Inches(2.5)) / 2
                ans = str(row['Answer'])
                parts = ans.split("-", 1) if "-" in ans else [ans, ""]
                add_styled_text(slide, parts[0].strip(), Inches(0.2), vt, cw - Inches(0.4), Inches(2.5), font_size=48,
                                bold=True, font_name="Gladiola", force_single_line=False)
                if path: place_smart_scaled_image(slide, path, prs.slide_width / 2, prs.slide_height / 2,
                                                  max_w=cw + Inches(1.5), max_h=Inches(5.8))
                add_styled_text(slide, parts[1].strip(), prs.slide_width - cw + Inches(0.2), vt, cw - Inches(0.4),
                                Inches(2.5), font_size=48, bold=True, font_name="Gladiola", force_single_line=False)
            elif r == 2:
                if path: place_smart_scaled_image(slide, path, prs.slide_width / 2, prs.slide_height / 2 - Inches(0.5),
                                                  max_w=Inches(8.0), max_h=Inches(4.5))
                add_styled_text(slide, f"Răspuns: {row['Answer']}", 0, Inches(5.8), prs.slide_width, Inches(1.5),
                                font_size=54, bold=True, font_name="Gladiola", force_single_line=False)
            else:
                add_styled_text(slide, row['Question'], Inches(0.5), Inches(2.2), Inches(6), Inches(2.5), is_qa=True,
                                font_name="Gladiola", force_single_line=False)
                add_styled_text(slide, f"Răspuns: {row['Answer']}", Inches(0.5), Inches(4.7), Inches(6), Inches(2.5),
                                font_size=44, bold=True, is_qa=True, font_name="Gladiola", force_single_line=False)
                if path: place_smart_scaled_image(slide, path, Inches(10), prs.slide_height / 2 + Inches(0.5),
                                                  max_w=Inches(5.0), max_h=Inches(4.5))
        if r in [3, 5]: add_styled_slide("Pauză", "Rezultatele momentului" if r == 3 else "Situația înainte de Pariu")

    p_data = df[df['Round'] == 6]
    if not p_data.empty:
        p_row = p_data.iloc[0]
        add_styled_slide("Pariul", "Miza crește...")
        sq = add_styled_slide()
        add_bg(sq)
        add_styled_text(sq, "Pariul: Întrebarea", Inches(0.5), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                        alignment=PP_ALIGN.LEFT, font_name="Gladiola", color_rgb=(0, 0, 255))
        w_img = glob.glob(os.path.join("images", "Wager", "*.*"))
        if w_img:
            add_styled_text(sq, p_row['Question'], Inches(0.5), Inches(1.0), Inches(6.0), Inches(6.0), bold=True,
                            is_qa=True, font_name="Gladiola", force_single_line=False)
            place_smart_scaled_image(sq, w_img[0], Inches(10), prs.slide_height / 2 + Inches(0.5), max_w=Inches(5.0),
                                     max_h=Inches(4.5))
        else:
            add_styled_text(sq, p_row['Question'], Inches(1.0), Inches(1.5), Inches(11.3), Inches(5.0), bold=True,
                            is_qa=True, font_name="Gladiola", force_single_line=False)
        sa = prs.slides.add_slide(prs.slide_layouts[6])
        add_bg(sa)
        add_styled_text(sa, "Pariul: Răspuns", Inches(0.5), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                        alignment=PP_ALIGN.LEFT, font_name="Gladiola", color_rgb=(0, 0, 255))
        add_styled_text(sa, f"Răspuns: {p_row['Answer']}", Inches(1.0), Inches(1.5), Inches(11.3), Inches(5.0),
                        font_size=54, bold=True, is_qa=True, font_name="Gladiola", force_single_line=False)
        add_styled_slide("Rezultate Finale", "Câștigătorii Transylvania Trivia")

    prs.save("Transylvania_Trivia.pptx")
    return "Transylvania_Trivia.pptx"


# --- 4. APP EXECUTION ---
tab_cmd, tab_score, tab_reveal = st.tabs(["🎮 Trivia Command", "🏆 Scoring Matrix", "📽️ Leaderboard Presentation"])
with tab_cmd:
    st.title("🧛 Trivia Production")
    uploaded = st.file_uploader("Load CSV", type="csv")
    if uploaded and st.button("Generate Slides"):
        path = generate_trivia_slides(pd.read_csv(uploaded))
        with open(path, "rb") as f: st.download_button("Download", f, file_name=path)
with tab_score: manage_scores()
with tab_reveal:
    if 'teams' in st.session_state and not st.session_state.teams.empty:
        lb = st.session_state.teams[['Echipă', 'Total']].sort_values(by='Total', ascending=True).reset_index(drop=True)
        if 'reveal_step' not in st.session_state: st.session_state.reveal_step = 0
        s = st.session_state.reveal_step
        if s < len(lb):
            curr = lb.iloc[s]
            rank = len(lb) - s
            nc = "#FFD700" if rank == 1 else ("#C0C0C0" if rank == 2 else ("#CD7F32" if rank == 3 else "#fff"))
            st.markdown(
                f'<div style="background-color:#000;padding:50px;text-align:center;border:8px solid #00004a;border-radius:20px;box-shadow:0 0 40px #000080;"><h2 style="color:#666;font-family:\'Chromium One\'">LOCUL</h2><h1 style="color:#fff;font-size:110px;font-family:\'Chromium One\'">{rank}</h1><h2 style="color:{nc};font-size:80px;">{curr["Echipă"]}</h2><h3 style="color:#fff;font-size:50px;">{curr["Total"]:.2f}</h3></div>',
                unsafe_allow_html=True)
            if st.button("Next"): st.session_state.reveal_step += 1; st.rerun()