import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
import os
import requests
import re
import io
import cv2
import copy
import time
from datetime import datetime
import streamlit as st

# --- STREAMLIT SETTINGS (Neon Green & Black Theme) ---
st.set_page_config(
    page_title="Auto Generate Report Observation", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS to match the "Vehicle Details" aesthetic
st.markdown("""
    <style>
    /* Main Background - Very Dark Gray/Black */
    .stApp {
        background-color: #0B0E0F;
    }
    
    /* Header Style */
    .custom-header {
        background-color: #111516;
        border-bottom: 2px solid #00FF66;
        padding: 15px;
        border-radius: 8px 8px 0px 0px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 20px;
    }
    
    .header-title {
        color: #00FF66 !important;
        font-family: 'Courier New', Courier, monospace;
        font-size: 22px;
        font-weight: bold;
        letter-spacing: 2px;
        margin: 0;
    }

    /* Labels - Small, Gray, Uppercase */
    label, .stMarkdown p {
        color: #888888 !important;
        font-family: 'Courier New', Courier, monospace !important;
        text-transform: uppercase;
        font-size: 12px !important;
        font-weight: bold;
    }

    /* Input Fields - Dark with Neon Green Text */
    .stTextInput>div>div>input, .stSelectbox>div>div>div, .stDateInput>div>div>input {
        background-color: #111516 !important;
        color: #00FF66 !important;
        border: 1px solid #333333 !important;
        font-family: 'Courier New', Courier, monospace !important;
        border-radius: 4px !important;
    }
    
    /* Success/Error Messages */
    .stAlert {
        background-color: #111516 !important;
        color: #00FF66 !important;
        border: 1px solid #00FF66 !important;
    }

    /* Buttons - Neon Green Outline/Solid */
    .stButton>button {
        background-color: transparent !important;
        color: #00FF66 !important;
        border: 2px solid #00FF66 !important;
        font-family: 'Courier New', Courier, monospace !important;
        font-weight: bold !important;
        transition: 0.3s;
    }
    
    .stButton>button:hover {
        background-color: #00FF66 !important;
        color: #000000 !important;
    }

    hr {
        border-color: #1A1E1F !important;
    }

    .stProgress > div > div > div > div {
        background-color: #00FF66 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# --- HEADER ---
st.markdown("""
    <div class="custom-header">
        <p class="header-title">OBSERVATION DETAILS</p>
        <span style="color: #FF3366; font-weight: bold; font-family: sans-serif; cursor: pointer;">✕</span>
    </div>
    """, unsafe_allow_html=True)

# --- REFRESH FUNCTION ---
def clear_all():
    for key in st.session_state.keys():
        del st.session_state[key]
    st.rerun()

# --- CORE LOGIC (Original preserved) ---

def download_and_insert_media(slide, link_data, left_inch, top_inch, width_inch, is_video_slide=False):
    if not isinstance(link_data, str) or "drive.google.com" not in link_data:
        return
    video_folder = "downloaded_videos"
    if not os.path.exists(video_folder): os.makedirs(video_folder)
    links = [l.strip() for l in link_data.split(';') if l.strip()]

    for index, clean_link in enumerate(links):
        file_id_match = re.search(r'[-\w]{25,}', clean_link)
        if file_id_match:
            file_id = file_id_match.group(0)
            session = requests.Session()
            download_url = "https://docs.google.com/uc?export=download"
            params = {'id': file_id, 'confirm': 't'}

            if is_video_slide:
                current_left = 2.0
                current_top = 1.5
                current_width = 6.0
            else:
                current_left = left_inch
                current_top = top_inch + (index * 2.1)
                current_width = width_inch

            try:
                response = session.get(download_url, params=params, stream=True, timeout=30)
                if response.status_code == 200:
                    content_type = response.headers.get('Content-Type', '')

                    if 'image' in content_type and not is_video_slide:
                        image_data = io.BytesIO(response.content)
                        slide.shapes.add_picture(image_data, Inches(current_left), Inches(current_top), width=Inches(current_width))

                    elif ('video' in content_type or 'octet-stream' in content_type) and is_video_slide:
                        video_path = os.path.join(video_folder, f"video_{file_id}.mp4")
                        thumb_path = os.path.join(video_folder, f"thumb_{file_id}.jpg")
                        with open(video_path, 'wb') as f:
                            for chunk in response.iter_content(chunk_size=8192): f.write(chunk)

                        vidcap = cv2.VideoCapture(video_path)
                        vidcap.set(cv2.CAP_PROP_POS_MSEC, 12000)
                        success, image = vidcap.read()
                        if not success:
                            vidcap.set(cv2.CAP_PROP_POS_MSEC, 0)
                            success, image = vidcap.read()
                        if success: cv2.imwrite(thumb_path, image)
                        vidcap.release()

                        slide.shapes.add_movie(video_path, Inches(current_left), Inches(current_top), width=Inches(current_width), height=Inches(current_width * 0.56), poster_frame_image=thumb_path if os.path.exists(thumb_path) else None, mime_type='video/mp4')
            except Exception as e: st.error(f"Error media {file_id}: {e}")

def create_custom_slide(pres, slide_template):
    blank_layout = pres.slide_layouts[6]
    new_slide = pres.slides.add_slide(blank_layout)

    for shp in list(new_slide.shapes):
        new_slide.shapes._spTree.remove(shp.element)

    slide_height = pres.slide_height
    footer_threshold = slide_height * 0.85

    for shape in slide_template.shapes:
        if shape.is_placeholder or shape.top > footer_threshold:
            continue

        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            img_stream = io.BytesIO(shape.image.blob)
            new_slide.shapes.add_picture(img_stream, shape.left, shape.top, shape.width, shape.height)
        elif shape.has_table:
            new_el = copy.deepcopy(shape.element)
            new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
        else:
            new_el = copy.deepcopy(shape.element)
            new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

    return new_slide

# --- STREAMLIT UI ---

# AUTOMATIC TEMPLATE CHECK
TEMPLATE_FILENAME = "template.pptx"
template_exists = os.path.exists(TEMPLATE_FILENAME)

if not template_exists:
    st.error(f"STATUS: '{TEMPLATE_FILENAME}' NOT FOUND")
else:
    st.success(f"STATUS: TEMPLATE LOADED")

# Inputs
col1, col2 = st.columns(2)

with col1:
    report_title = st.text_input("REPORT TITLE", placeholder="Input Title...", key="ti_title")
    obs_month = st.selectbox("OBSERVATION MONTH", ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"], key="sb_month")
    start_date = st.date_input("START DATE", value=datetime(2026, 1, 1), key="di_start")

with col2:
    depot_list = ["MRT Jinjang", "Shah Alam", "Cheras Selatan", "Batu Caves", "MRT Kajang", "MRT Sungai Buloh", "MRT Serdang"]
    selected_depot = st.selectbox("DEPOT LOCATION", depot_list, key="sb_depot")
    siri_list = [f"OE/SQI/CR/VO/{str(i).zfill(3)}/2026" for i in range(1, 101)]
    selected_siri = st.selectbox("SIRI NUMBER", siri_list, key="sb_siri")
    end_date = st.date_input("END DATE", value=datetime(2026, 12, 31), key="di_end")

st.divider()

# File Uploader
uploaded_excel = st.file_uploader("UPLOAD DATA SOURCE (EXCEL)", type=["xlsx", "xls"], key="fu_excel")

# Action Buttons Side by Side
btn_col1, btn_col2 = st.columns([3, 1])

with btn_col1:
    generate_btn = st.button("RUN GENERATOR", use_container_width=True)

with btn_col2:
    # This button now calls the clear_all function directly
    if st.button("REFRESH", use_container_width=True, on_click=clear_all):
        pass

# --- EXECUTION LOGIC ---
if generate_btn:
    if not uploaded_excel:
        st.error("ERROR: EXCEL FILE MISSING")
    elif not template_exists:
        st.error("ERROR: TEMPLATE MISSING")
    else:
        try:
            start_time = time.time()
            df = pd.read_excel(uploaded_excel, sheet_name='Sheet1')
            df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')

            def safe_date_convert(x):
                try: return pd.to_datetime(x).date()
                except: return None

            df['cleaned_date'] = df.iloc[:, 0].apply(safe_date_convert)
            mask = (df['cleaned_date'] >= start_date) & \
                   (df['cleaned_date'] <= end_date) & \
                   (df.iloc[:, 3].astype(str).str.strip() == selected_depot)

            filtered_data = df.loc[mask].copy()

            if filtered_data.empty:
                st.warning(f"NO RECORDS FOUND FOR {selected_depot}")
            else:
                prs = Presentation(TEMPLATE_FILENAME)
                slide6_template = prs.slides[5] if len(prs.slides) >= 6 else None
                slide1_template = prs.slides[0]
                slide2_template = prs.slides[1]
                slide3_template = prs.slides[2]

                processed_count = 0
                total = len(filtered_data)
                progress_bar = st.progress(0)
                status_text = st.empty()
                summary_list = []

                for _, row in filtered_data.iterrows():
                    if str(row.iloc[8]).strip().lower() == "yes": continue

                    summary_list.append(row)
                    new_data_slide = create_custom_slide(prs, slide1_template)
                    dt_raw = pd.to_datetime(row.iloc[0])
                    date_str = dt_raw.strftime('%d/%m/%Y') if not pd.isnull(dt_raw) else "N/A"
                    time_str = dt_raw.strftime('%H:%M:%S') if not pd.isnull(dt_raw) else "N/A"

                    col_i, col_j = str(row.iloc[8]).lower(), str(row.iloc[9]).lower()
                    pemerhatian = ""
                    if col_i == "no": pemerhatian += "1. Pelanggaran Had Laju Hentian: BC memintas hentian dengan kelajuan melebihi 25 km/j.\n"
                    if col_j == "no": pemerhatian += "2. Kapten Bas tidak memandu / menggunakan lorong kiri"

                    cadangan = ""
                    if col_i == "no" and col_j == "no": cadangan = "1. Memberi peringatan/kaunseling kepada Kapten Bas memperlahankan bas and keperluan berada di lorong kiri."
                    elif col_i == "no": cadangan = "1. Memberi peringatan/kaunseling kepada Kapten Bas memperlahankan bas di setiap hentian bas."
                    elif col_j == "no": cadangan = "2. Memberi peringatan kepada Kapten Bas mengenai keperluan berada di lorong kiri."

                    replacements = {
                        "Tarikh pemerhatian :": f"Tarikh pemerhatian : {date_str}",
                        "Nombor Bas :": f"Nombor Bas : {row.iloc[6]}",
                        "Laluan pemerhatian :": f"Laluan pemerhatian : {row.iloc[4]}",
                        "Masa :": f"Masa : {time_str}",
                        "Lokasi / Hentian :": f"Lokasi / Hentian : {row.iloc[5]}",
                        "Nama Kapten Bas :": f"Nama Kapten Bas : {row.iloc[32]}",
                        "ID Kapten Bas :": f"ID Kapten Bas : {row.iloc[31]}",
                        "Kelajuan Dipandu :": f"Kelajuan Dipandu : {row.iloc[30]} Km/h",
                        "Nama PIC :": f"Nama PIC : {row.iloc[2]}",
                        "Pemerhatian Pemanduan Kapten Bas :": f"Pemerhatian Pemanduan Kapten Bas :\n{pemerhatian}",
                        "Cadangan:": f"Cadangan:\n{cadangan}"
                    }

                    for shape in new_data_slide.shapes:
                        if shape.has_table:
                            for r in shape.table.rows:
                                for cell in r.cells:
                                    for paragraph in cell.text_frame.paragraphs:
                                        for key, value in replacements.items():
                                            if key in paragraph.text:
                                                paragraph.text = paragraph.text.replace(key, str(value))
                                                for run in paragraph.runs: run.font.size = Pt(10)

                    download_and_insert_media(new_data_slide, row.iloc[26], left_inch=0.6, top_inch=2.1, width_inch=3.8, is_video_slide=False)
                    new_video_slide = create_custom_slide(prs, slide2_template)
                    download_and_insert_media(new_video_slide, row.iloc[26], left_inch=0, top_inch=0, width_inch=0, is_video_slide=True)

                    processed_count += 1
                    progress_bar.progress(processed_count / total)
                    status_text.text(f"COMPILING... {processed_count}/{total}")

                if summary_list:
                    new_summary_slide = create_custom_slide(prs, slide3_template)
                    orig_table_shape = next((s for s in new_summary_slide.shapes if s.has_table), None)
                    if orig_table_shape:
                        height = orig_table_shape.height
                        rows_needed, cols_needed = len(summary_list) + 1, 6
                        new_summary_slide.shapes._spTree.remove(orig_table_shape.element)
                        new_table_shape = new_summary_slide.shapes.add_table(rows_needed, cols_needed, Inches(0.5), Inches(1.5), Inches(9.0), height)
                        summary_table = new_table_shape.table
                        summary_table.columns[0].width = Inches(1.2); summary_table.columns[1].width = Inches(0.8)
                        summary_table.columns[2].width = Inches(1.0); summary_table.columns[3].width = Inches(2.2)
                        summary_table.columns[4].width = Inches(2.6); summary_table.columns[5].width = Inches(1.2)

                        headers = ["Depoh", "Laluan", "Nombor Bas", "Hentian Bas", "Pengesahan", "Status"]
                        for i, h_text in enumerate(headers):
                            cell = summary_table.rows[0].cells[i]
                            cell.text = h_text
                            for para in cell.text_frame.paragraphs:
                                for run in para.runs: run.font.size = Pt(10); run.font.bold = True

                        for idx, s_row in enumerate(summary_list):
                            tr = summary_table.rows[idx + 1]
                            tr.height = Inches(0.7)
                            tr.cells[0].text = str(s_row.iloc[3]); tr.cells[1].text = str(s_row.iloc[4]); tr.cells[2].text = str(s_row.iloc[6]); tr.cells[3].text = str(s_row.iloc[5])
                            dt_full = pd.to_datetime(s_row.iloc[0]).strftime('%d/%m/%Y %H:%M:%S')
                            tr.cells[4].text = f"ID: {s_row.iloc[31]}\nNama: {s_row.iloc[32]}\nLaju: {s_row.iloc[30]} Km/h\nMasa: {dt_full}"
                            tr.cells[5].text = "Tidak Mematuhi"
                            for cell in tr.cells:
                                for para in cell.text_frame.paragraphs:
                                    for run in para.runs: run.font.size = Pt(8); run.font.name = "Arial"

                if slide6_template: create_custom_slide(prs, slide6_template)

                # Reordering & Replacing Title
                xml_slides = prs.slides._sldIdLst
                for _ in range(3): xml_slides.remove(xml_slides[0])
                if len(prs.slides) >= 3:
                    summary_slide_element = xml_slides[len(xml_slides) - 2]
                    xml_slides.remove(summary_slide_element); xml_slides.insert(2, summary_slide_element)
                    slide_two_element = xml_slides[1]
                    xml_slides.remove(slide_two_element); xml_slides.insert(0, slide_two_element)
                if len(xml_slides) >= 4: xml_slides.remove(xml_slides[3])

                full_replacement_text = f"Central Region {obs_month} {datetime.now().year}"
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            if "Januari-February 2026" in shape.text_frame.text:
                                for paragraph in shape.text_frame.paragraphs:
                                    if "Januari-February 2026" in paragraph.text:
                                        paragraph.text = paragraph.text.replace("Januari-February 2026", full_replacement_text)
                                        for run in paragraph.runs: run.font.name = 'Arial'; run.font.size = Pt(20); run.font.color.rgb = RGBColor(0, 32, 96)
                            if "OE/SQI/CR/VO/001/2026" in shape.text_frame.text:
                                for paragraph in shape.text_frame.paragraphs:
                                    if "OE/SQI/CR/VO/001/2026" in paragraph.text:
                                        paragraph.text = paragraph.text.replace("OE/SQI/CR/VO/001/2026", selected_siri)
                                        for run in paragraph.runs: run.font.size = Pt(12)

                ppt_output = io.BytesIO()
                prs.save(ppt_output)
                ppt_output.seek(0)
                
                duration = time.time() - start_time
                st.success(f"COMPLETE: {int(duration // 60)}M {int(duration % 60)}S")
                
                st.download_button(
                    label="DOWNLOAD GENERATED PPTX",
                    data=ppt_output,
                    file_name=f"{report_title if report_title else 'Report'}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"SYSTEM ERROR: {str(e)}")
