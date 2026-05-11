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

# --- STREAMLIT SETTINGS (Dark Theme) ---
st.set_page_config(
    page_title="Auto Generate Report Observation", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for a deeper "Black Theme" look
st.markdown("""
    <style>
    .stApp {
        background-color: #0E1117;
        color: #FFFFFF;
    }
    .stButton>button {
        border-radius: 5px;
    }
    </style>
    """, unsafe_allow_html=True)

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

st.title("📊 Auto Generate Report Observation")
st.subheader("Sure Stop v1")

# AUTOMATIC TEMPLATE CHECK
TEMPLATE_FILENAME = "template.pptx"
template_exists = os.path.exists(TEMPLATE_FILENAME)

if not template_exists:
    st.error(f"❌ '{TEMPLATE_FILENAME}' not found in GitHub repository. Please upload it to your repo.")
else:
    st.success(f"✅ Template '{TEMPLATE_FILENAME}' loaded from repository.")

# Inputs
col1, col2 = st.columns(2)

with col1:
    report_title = st.text_input("Title", placeholder="e.g., Monthly Observation")
    obs_month = st.selectbox("Observation Month", ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])
    start_date = st.date_input("Start Date", value=datetime(2026, 1, 1))

with col2:
    depot_list = ["MRT Jinjang", "Shah Alam", "Cheras Selatan", "Batu Caves", "MRT Kajang", "MRT Sungai Buloh", "MRT Serdang"]
    selected_depot = st.selectbox("Depot", depot_list)
    siri_list = [f"OE/SQI/CR/VO/{str(i).zfill(3)}/2026" for i in range(1, 101)]
    selected_siri = st.selectbox("No. Siri", siri_list)
    end_date = st.date_input("End Date", value=datetime(2026, 12, 31))

st.divider()

# File Uploader for Excel
uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

# Action Buttons Side by Side
btn_col1, btn_col2 = st.columns([3, 1])

with btn_col1:
    generate_btn = st.button("🚀 Generate Report", use_container_width=True)

with btn_col2:
    if st.button("🔄 Refresh", use_container_width=True):
        st.rerun()

if generate_btn:
    if not uploaded_excel:
        st.error("Please upload the Excel file.")
    elif not template_exists:
        st.error("Cannot proceed: Template file is missing from the repository.")
    else:
        try:
            start_time = time.time()
            
            # Load Data
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
                st.warning(f"No records found for {selected_depot} in the selected date range.")
            else:
                # Load Presentation from local file (GitHub Repo)
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
                    status_text.text(f"Processing... {processed_count}/{total}")

                if summary_list:
                    new_summary_slide = create_custom_slide(prs, slide3_template)
                    orig_table_shape = next((s for s in new_summary_slide.shapes if s.has_table), None)

                    if orig_table_shape:
                        height = orig_table_shape.height
                        rows_needed, cols_needed = len(summary_list) + 1, 6
                        style_id = orig_table_shape.table._tbl.find('.//a:tableStyleId', namespaces=orig_table_shape.table._tbl.nsmap)
                        style_id_val = style_id.text if style_id is not None else "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

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
                                for run in para.runs:
                                    run.font.size = Pt(10); run.font.bold = True

                        for idx, s_row in enumerate(summary_list):
                            tr = summary_table.rows[idx + 1]
                            tr.height = Inches(0.7)
                            tr.cells[0].text = str(s_row.iloc[3])
                            tr.cells[1].text = str(s_row.iloc[4])
                            tr.cells[2].text = str(s_row.iloc[6])
                            tr.cells[3].text = str(s_row.iloc[5])
                            dt_full = pd.to_datetime(s_row.iloc[0]).strftime('%d/%m/%Y %H:%M:%S')
                            tr.cells[4].text = f"ID: {s_row.iloc[31]}\nNama: {s_row.iloc[32]}\nLaju: {s_row.iloc[30]} Km/h\nMasa: {dt_full}"
                            tr.cells[5].text = "Tidak Mematuhi"
                            for cell in tr.cells:
                                for para in cell.text_frame.paragraphs:
                                    for run in para.runs: run.font.size = Pt(8); run.font.name = "Arial"

                if slide6_template:
                    create_custom_slide(prs, slide6_template)

                # Slide Reordering Logic
                xml_slides = prs.slides._sldIdLst
                for _ in range(3): xml_slides.remove(xml_slides[0])
                if len(prs.slides) >= 3:
                    summary_slide_element = xml_slides[len(xml_slides) - 2]
                    xml_slides.remove(summary_slide_element)
                    xml_slides.insert(2, summary_slide_element)
                    slide_two_element = xml_slides[1]
                    xml_slides.remove(slide_two_element)
                    xml_slides.insert(0, slide_two_element)
                if len(xml_slides) >= 4: xml_slides.remove(xml_slides[3])

                # Text replacements
                current_year = datetime.now().year
                full_replacement_text = f"Central Region {obs_month} {current_year}"

                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            if "Januari-February 2026" in shape.text_frame.text:
                                for paragraph in shape.text_frame.paragraphs:
                                    if "Januari-February 2026" in paragraph.text:
                                        paragraph.text = paragraph.text.replace("Januari-February 2026", full_replacement_text)
                                        for run in paragraph.runs:
                                            run.font.name = 'Arial'; run.font.size = Pt(20)
                                            run.font.color.rgb = RGBColor(0, 32, 96)
                            
                            if "OE/SQI/CR/VO/001/2026" in shape.text_frame.text:
                                for paragraph in shape.text_frame.paragraphs:
                                    if "OE/SQI/CR/VO/001/2026" in paragraph.text:
                                        paragraph.text = paragraph.text.replace("OE/SQI/CR/VO/001/2026", selected_siri)
                                        for run in paragraph.runs: run.font.size = Pt(12)

                # Finalize
                ppt_output = io.BytesIO()
                prs.save(ppt_output)
                ppt_output.seek(0)
                
                duration = time.time() - start_time
                st.success(f"Done! Processed in {int(duration // 60)}m {int(duration % 60)}s")
                
                file_name = f"{report_title if report_title else 'Generated_Report'}.pptx"
                st.download_button(
                    label="📥 Download Presentation",
                    data=ppt_output,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
