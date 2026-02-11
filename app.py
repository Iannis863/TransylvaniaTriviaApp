import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import os
import glob
import io
import re

# --- PAGE CONFIG ---
st.set_page_config(page_title="Transylvania Trivia Command", layout="wide", page_icon="🧛")

# --- FOLDER & MASTER FILE SETUP ---
GRADING_FOLDER = "Grading Sheets"
MASTER_FILE_PATH = os.path.join(GRADING_FOLDER, "Grading Sheet Overall.xlsx")

if not os.path.exists(GRADING_FOLDER):
    os.makedirs(GRADING_FOLDER)


# --- HELPER FOR ORDINAL SUFFIX ---
def get_ordinal(n):
    if 11 <= (n % 100) <= 13:
        return f"{n}th"
    return f"{n}" + {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')


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
                new_t = pd.DataFrame([{
                    'Echipă': t_name, 'R1': 0.00, 'R2': 0.00, 'R3': 0.00, 'R4': 0.00, 'R5': 0.00,
                    'Joker': "Niciuna", 'Pariu (±)': 0.00, 'Total': 0.00
                }])
                st.session_state.teams = pd.concat([st.session_state.teams, new_t], ignore_index=True)
                st.rerun()

    st.markdown("---")
    col_input, col_reset, col_push = st.columns([1, 1, 1])

    with col_input:
        edition = st.number_input("Ediția Trivia (ex: 5)", min_value=1, step=1, value=1, label_visibility="collapsed")

    with col_reset:
        if st.button("🧹 Reset / Șterge Tot", use_container_width=True):
            st.session_state.teams = pd.DataFrame(columns=COLS)
            if 'reveal_step' in st.session_state:
                st.session_state.reveal_step = 0
            st.rerun()

    with col_push:
        if not st.session_state.teams.empty:
            edition_str = get_ordinal(edition)
            session_filename = f"Grading Sheet {edition_str} edition.xlsx"
            session_path = os.path.join(GRADING_FOLDER, session_filename)

            if st.button(f"🚀 Push to Ed. {edition}", use_container_width=True):
                # 1. SAVE INDIVIDUAL SESSION SHEET
                output_session = io.BytesIO()
                with pd.ExcelWriter(output_session, engine='xlsxwriter') as writer:
                    export_df = st.session_state.teams.sort_values(by='Total', ascending=False)
                    export_df.to_excel(writer, index=False, sheet_name='Final Scores')
                    workbook = writer.book
                    worksheet = writer.sheets['Final Scores']
                    header_fmt = workbook.add_format(
                        {'bold': True, 'bg_color': '#4A0000', 'font_color': 'white', 'border': 1})
                    for col_num, value in enumerate(export_df.columns.values):
                        worksheet.write(0, col_num, value, header_fmt)
                        worksheet.set_column(col_num, col_num, 18)

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

                if edition_col in df_master.columns:
                    df_master = df_master.drop(columns=[edition_col])

                df_master = pd.merge(df_master, current_scores, on='Numele echipei', how='outer')

                # SORT COLUMNS: Team Name, Editiile (Numerical), Final Score
                cols = df_master.columns.tolist()
                edition_cols = [c for c in cols if "Ediția" in c]
                edition_cols.sort(key=lambda x: int(re.search(r'\d+', x).group()))
                final_col_order = ['Numele echipei'] + edition_cols + ['Scor Final']
                df_master = df_master[final_col_order]

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
                        workbook = writer.book
                        worksheet = writer.sheets['Overall Ranking']
                        header_fmt = workbook.add_format(
                            {'bold': True, 'bg_color': '#00004A', 'font_color': 'white', 'border': 1})
                        for col_num, value in enumerate(df_master.columns.values):
                            worksheet.write(0, col_num, value, header_fmt)
                            worksheet.set_column(col_num, col_num, 22)
                    st.toast(f"✅ Master Sheet Updated", icon='🧛')
                except PermissionError:
                    st.error(f"❌ ACCES REFUZAT: Închide '{os.path.basename(MASTER_FILE_PATH)}' și apasă PUSH din nou!")

    def run_calculation(df):
        if df.empty: return df
        numeric_cols = ['R1', 'R2', 'R3', 'R4', 'R5', 'Pariu (±)']
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.00)

        def row_math(row):
            rounds = ["R1", "R2", "R3", "R4", "R5"]
            score_sum = 0.00
            j_choice = str(row['Joker'])
            for r in rounds:
                val = float(row[r])
                score_sum += (val * 2) if j_choice == r else val
            return score_sum + float(row['Pariu (±)'])

        df['Total'] = df.apply(row_math, axis=1)
        return df

    st.markdown("### 📊 Tabel Scoruri")
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
        hide_index=False, use_container_width=True, num_rows="dynamic", key="score_editor"
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
    if 'reveal_step' not in st.session_state:
        st.session_state.reveal_step = 0
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
            .stButton>button[kind="secondary"]:nth-of-type(2) {{
                width: 100%; height: 55vh; background-color: transparent !important;
                border: none !important; color: transparent !important; z-index: 999; position: absolute;
            }}
            .reveal-container {{
                background-color: #000000; padding: 50px 80px; border-radius: 20px; border: 8px solid #4a0000;
                text-align: center; min-height: 55vh; display: flex; flex-direction: column; justify-content: center;
                align-items: center; box-shadow: 0 0 40px #800000;
            }}
        </style>
        <div class="reveal-container">
            <h2 style="color: #666666; font-size: 30px; margin: 0; font-family: 'Chromium One', sans-serif;">LOCUL</h2>
            <h1 style="color: #ffffff; font-size: 110px; margin: 0; line-height: 1; font-family: 'Chromium One', sans-serif;">{rank}</h1>
            <div style="width: 250px; height: 3px; background-color: #800000; margin: 20px 0;"></div>
            <h2 style="color: {name_color}; font-size: 80px; font-weight: bold; margin: 5px 0;">{current_team['Echipă']}</h2>
            <h3 style="color: #ffffff; font-size: 50px; margin: 0;">{current_team['Total']:.2f}</h3>
        </div>
    """, unsafe_allow_html=True)
    if st.button("Următorul Loc", key="screen_click_trigger"):
        st.session_state.reveal_step += 1
        st.rerun()


# --- 3. STYLE HELPERS ---
def get_dynamic_font_size(text, is_qa=False):
    length = len(str(text))
    if is_qa: return 64 if length < 15 else (52 if length < 30 else 42)
    return 58 if length < 30 else (48 if length < 70 else 38)


def add_background_to_slide(slide, prs, specific_path=None):
    bg_path = specific_path if specific_path else os.path.join("images", "background.jpeg")
    if os.path.exists(bg_path):
        pic = slide.shapes.add_picture(bg_path, 0, 0, width=prs.slide_width, height=prs.slide_height)
        shape_tree = slide.shapes._spTree
        shape_tree.remove(pic._element)
        shape_tree.insert(2, pic._element)


def place_smart_scaled_image(slide, img_path, target_center_x, target_center_y, max_w=Inches(5.5), max_h=Inches(4.5)):
    if not os.path.exists(img_path): return
    pic = slide.shapes.add_picture(img_path, 0, 0)
    ratio = min(max_w / pic.width, max_h / pic.height)
    new_w, new_h = int(pic.width * ratio), int(pic.height * ratio)
    pic.width, pic.height = new_w, new_h
    pic.left, pic.top = int(target_center_x - (new_w / 2)), int(target_center_y - (new_h / 2))


def add_styled_text(slide, text, left, top, width, height, font_size=None, is_qa=False, bold=False,
                    alignment=PP_ALIGN.CENTER, font_name=None, color_rgb=(255, 255, 255), force_single_line=True):
    if not text or str(text).strip() == "": return
    if font_size is None: font_size = get_dynamic_font_size(text, is_qa)
    txBox = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    tf = txBox.text_frame
    tf.word_wrap = not force_single_line
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = str(text)
    p.alignment = alignment
    if len(p.runs) > 0:
        run = p.runs[0]
        if force_single_line:
            ratio = 1.4 if font_name == "Chromium One" else 1.9
            current_size = font_size
            while (len(str(text)) * (current_size / ratio)) > (width / 914400 * 72):
                current_size -= 5
                if current_size < 20: break
            font_size = current_size
        run.font.size, run.font.bold = Pt(font_size), bold
        if font_name: run.font.name = font_name
        run.font.color.rgb = RGBColor(*color_rgb)


# --- 4. CORE GENERATOR ---
def generate_trivia_slides(df):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    intro_images = ["First Slide.png", "Second Slide.png", "Third Slide.png"]
    for img_name in intro_images:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        full_img_path = os.path.join("images", "Slides", img_name)
        add_background_to_slide(slide, prs, specific_path=full_img_path)

    def add_styled_slide(title="", subtitle=""):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_background_to_slide(slide, prs)
        if title: add_styled_text(slide, title, Inches(0.5), Inches(0.5), prs.slide_width - Inches(1.0), Inches(3.0),
                                  font_size=150, font_name="Chromium One", color_rgb=(0, 0, 255),
                                  force_single_line=True)
        if subtitle: add_styled_text(slide, subtitle, Inches(0.5), Inches(4.0), prs.slide_width - Inches(1.0),
                                     Inches(3.0), font_size=125, font_name="Gladiola", color_rgb=(173, 216, 230),
                                     force_single_line=True)
        return slide

    rounds = sorted([r for r in df['Round'].unique() if r <= 5])
    for r in rounds:
        r_data = df[df['Round'] == r]
        base_r_name = r_data['Round_Name'].iloc[0]
        add_styled_slide(f"Runda {r}", base_r_name)
        for i, (_, row) in enumerate(r_data.iterrows(), 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_background_to_slide(slide, prs)
            current_display_name = row['Round_Name'] if r == 2 else base_r_name
            add_styled_text(slide, current_display_name, Inches(0.5), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                            alignment=PP_ALIGN.LEFT, font_name="Gladiola", color_rgb=(0, 0, 255))
            add_styled_text(slide, f"Întrebarea {i}", Inches(7.8), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                            alignment=PP_ALIGN.RIGHT, font_name="Gladiola", color_rgb=(0, 0, 255))
            img_path = os.path.join("images", f"Round {r}", str(row.get('Picture'))) if pd.notna(
                row.get('Picture')) else None
            if r == 3:
                audio_path = os.path.join("audio", str(row['Question']))
                if os.path.exists(audio_path):
                    movie = slide.shapes.add_movie(os.path.abspath(audio_path), Inches(-1), Inches(-1), Inches(0.5),
                                                   Inches(0.5))
                    movie.name = "TriviaAudio"
            if r in [2, 3]:
                if img_path: place_smart_scaled_image(slide, img_path, prs.slide_width / 2,
                                                      prs.slide_height / 2 + Inches(0.4), max_w=Inches(9.0),
                                                      max_h=Inches(5.0))
            else:
                add_styled_text(slide, row['Question'], Inches(0.5), Inches(1.0), Inches(6.0), Inches(6.0), bold=True,
                                is_qa=True, force_single_line=False)
                if img_path: place_smart_scaled_image(slide, img_path, Inches(10), prs.slide_height / 2 + Inches(0.5),
                                                      max_w=Inches(5.0), max_h=Inches(4.5))

        add_styled_slide(f"Runda {r}", "Răspunsuri")
        for i, (_, row) in enumerate(r_data.iterrows(), 1):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            add_background_to_slide(slide, prs)
            current_display_name = row['Round_Name'] if r == 2 else base_r_name
            add_styled_text(slide, current_display_name, Inches(0.5), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                            alignment=PP_ALIGN.LEFT, font_name="Gladiola", color_rgb=(0, 0, 255))
            add_styled_text(slide, f"Răspuns {i}", Inches(7.8), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                            alignment=PP_ALIGN.RIGHT, font_name="Gladiola", color_rgb=(0, 0, 255))
            img_path = os.path.join("images", f"Round {r}", str(row.get('Picture'))) if pd.notna(
                row.get('Picture')) else None

            if r == 3:
                # Column Split Logic: Slide Width is 13.33"
                col_width = prs.slide_width / 3  # ~4.44 inches
                text_height = Inches(2.5)
                v_top = (prs.slide_height - text_height) / 2

                # Split Answer
                full_ans = str(row['Answer'])
                if "-" in full_ans:
                    parts = full_ans.split("-", 1)
                    song, artist = parts[0].strip(), parts[1].strip()
                else:
                    song, artist = full_ans.strip(), ""

                # 1. SONG (LEFT ZONE)
                add_styled_text(slide, song, Inches(0.2), v_top, col_width - Inches(0.4), text_height, font_size=48,
                                bold=True, is_qa=True, alignment=PP_ALIGN.CENTER, font_name="Gladiola",
                                force_single_line=False)

                # 2. PICTURE (MIDDLE ZONE)
                # Bigger scaling within the middle area
                if img_path: place_smart_scaled_image(slide, img_path, prs.slide_width / 2, prs.slide_height / 2,
                                                      max_w=col_width + Inches(1.5), max_h=Inches(5.8))

                # 3. ARTIST (RIGHT ZONE)
                add_styled_text(slide, artist, prs.slide_width - col_width + Inches(0.2), v_top,
                                col_width - Inches(0.4), text_height, font_size=48, bold=True, is_qa=True,
                                alignment=PP_ALIGN.CENTER, font_name="Gladiola", force_single_line=False)

            elif r in [2]:
                if img_path: place_smart_scaled_image(slide, img_path, prs.slide_width / 2,
                                                      prs.slide_height / 2 - Inches(0.5), max_w=Inches(8.0),
                                                      max_h=Inches(4.5))
                add_styled_text(slide, f"Răspuns: {row['Answer']}", 0, Inches(5.8), prs.slide_width, Inches(1.5),
                                font_size=54, bold=True, is_qa=True, force_single_line=False)
            else:
                add_styled_text(slide, row['Question'], Inches(0.5), Inches(2.2), Inches(6.0), Inches(2.5), is_qa=True,
                                force_single_line=False)
                add_styled_text(slide, f"Răspuns: {row['Answer']}", Inches(0.5), Inches(4.7), Inches(6.0), Inches(2.5),
                                font_size=44, bold=True, is_qa=True, force_single_line=False)
                if img_path: place_smart_scaled_image(slide, img_path, Inches(10), prs.slide_height / 2 + Inches(0.5),
                                                      max_w=Inches(5.0), max_h=Inches(4.5))

        if r in [3, 5]: add_styled_slide("Pauză", "Rezultatele momentului" if r == 3 else "Situația înainte de Pariu")

    p_data = df[df['Round'] == 6]
    if not p_data.empty:
        p_row = p_data.iloc[0]
        add_styled_slide("Pariul", "Miza crește...")
        sq = add_styled_slide()
        add_styled_text(sq, "Pariul: Întrebarea", Inches(0.5), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                        alignment=PP_ALIGN.LEFT, font_name="Gladiola", color_rgb=(0, 0, 255))
        wager_img = glob.glob(os.path.join("images", "Wager", "*.*"))
        if wager_img:
            add_styled_text(sq, p_row['Question'], Inches(0.5), Inches(1.0), Inches(6.0), Inches(6.0), bold=True,
                            is_qa=True, force_single_line=False)
            place_smart_scaled_image(sq, wager_img[0], Inches(10), prs.slide_height / 2 + Inches(0.5),
                                     max_w=Inches(5.0), max_h=Inches(4.5))
        else:
            add_styled_text(sq, p_row['Question'], Inches(1.0), Inches(1.5), Inches(11.3), Inches(5.0), bold=True,
                            is_qa=True, force_single_line=False)
        add_styled_slide("Răspunsul", "Momentul adevărului")
        sa = prs.slides.add_slide(prs.slide_layouts[6])
        add_background_to_slide(sa, prs)
        add_styled_text(sa, "Pariul: Răspuns", Inches(0.5), Inches(0.2), Inches(5), Inches(0.8), font_size=44,
                        alignment=PP_ALIGN.LEFT, font_name="Gladiola", color_rgb=(0, 0, 255))
        add_styled_text(sa, f"Răspuns: {p_row['Answer']}", Inches(1.0), Inches(1.5), Inches(11.3), Inches(5.0),
                        font_size=54, bold=True, is_qa=True, force_single_line=False)
        add_styled_slide("Rezultate Finale", "Câștigătorii Transylvania Trivia")

    filename = "Transylvania_Trivia.pptx"
    prs.save(filename)
    return filename


# --- 5. APP EXECUTION ---
tab_cmd, tab_score, tab_reveal = st.tabs(["🎮 Trivia Command", "🏆 Scoring Matrix", "📽️ Leaderboard Presentation"])
with tab_cmd:
    st.title("🧛 Trivia Production")
    uploaded = st.file_uploader("Load CSV file", type="csv")
    if uploaded and st.button("Generate Slides"):
        path = generate_trivia_slides(pd.read_csv(uploaded))
        with open(path, "rb") as f: st.download_button("Download", f, file_name=path)
with tab_score: manage_scores()
with tab_reveal: show_leaderboard_reveal()