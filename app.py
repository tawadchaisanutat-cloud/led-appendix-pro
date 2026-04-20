import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import re
import requests
import fitz  # PyMuPDF
from PIL import Image

# --- การตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="LED Appendix Pro v3 (16:9)", layout="wide")

# --- 1. ระบบจัดการสถานะและ Cache ---
if 'selected_slides' not in st.session_state:
    st.session_state.selected_slides = set()
if 'selected_drive_images' not in st.session_state:
    st.session_state.selected_drive_images = set()
if 'drive_files' not in st.session_state:
    st.session_state.drive_files = []

@st.cache_resource
def process_file_optimized(file_content, file_name):
    f_ext = file_name.split('.')[-1].lower()
    slides_data = []

    if f_ext == "pptx":
        prs = Presentation(io.BytesIO(file_content))
        for i, slide in enumerate(prs.slides):
            preview = None
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    preview = io.BytesIO(shape.image.blob)
                    break
            slides_data.append({"id": i, "type": "pptx", "display": f"Slide {i+1}", "preview": preview})
        return prs, slides_data

    elif f_ext == "pdf":
        doc = fitz.open(stream=file_content, filetype="pdf")
        for i in range(len(doc)):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(0.8, 0.8))
            slides_data.append({
                "id": i, "type": "pdf", "display": f"หน้า {i+1}",
                "preview": io.BytesIO(pix.tobytes("png"))
            })
        return doc, slides_data

# --- Google Drive helpers ---
def extract_folder_id(url):
    for pattern in [r'folders/([a-zA-Z0-9_-]+)', r'id=([a-zA-Z0-9_-]+)']:
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return None

@st.cache_data(show_spinner=False)
def list_drive_images(api_key, folder_id):
    params = {
        "q": f"'{folder_id}' in parents and mimeType contains 'image/' and trashed=false",
        "key": api_key,
        "fields": "files(id,name,mimeType)",
        "pageSize": 200,
    }
    r = requests.get("https://www.googleapis.com/drive/v3/files", params=params, timeout=15)
    r.raise_for_status()
    return r.json().get("files", [])

@st.cache_data(show_spinner=False)
def download_drive_image(file_id, api_key):
    r = requests.get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",
        params={"alt": "media", "key": api_key},
        timeout=30,
    )
    r.raise_for_status()
    return r.content

# --- 2. ฟังก์ชันช่วยจัดการ Layout ---
def copy_text_with_format(src_shape, target_slide):
    try:
        new_shape = target_slide.shapes.add_textbox(src_shape.left, src_shape.top, src_shape.width, src_shape.height)
        new_tf = new_shape.text_frame
        new_tf.word_wrap = src_shape.text_frame.word_wrap
        for i, src_para in enumerate(src_shape.text_frame.paragraphs):
            new_p = new_tf.paragraphs[0] if i == 0 else new_tf.add_paragraph()
            new_p.alignment = src_para.alignment
            for src_run in src_para.runs:
                nr = new_p.add_run()
                nr.text = src_run.text
                if src_run.font.size: nr.font.size = src_run.font.size
                if src_run.font.bold: nr.font.bold = src_run.font.bold
    except: pass

def add_full_image_16_9(slide, image_stream, prs):
    img = Image.open(image_stream)
    img_w, img_h = img.size
    ratio = min(prs.slide_width / img_w, prs.slide_height / img_h)
    new_w, new_h = int(img_w * ratio), int(img_h * ratio)
    left = (prs.slide_width - new_w) // 2
    top = (prs.slide_height - new_h) // 2
    image_stream.seek(0)
    slide.shapes.add_picture(image_stream, left, top, width=new_w, height=new_h)

# --- 3. ส่วน UI และ Logic หลัก ---
st.title("🚀 Professional LED Appendix Builder (16:9)")

with st.sidebar:
    st.header("1. ตั้งค่าไฟล์")
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Master (PPTX/PDF)", type=["pptx", "pdf"])
    cover_title = st.text_input("หัวข้อหน้าปก", "PROPOSAL FOR DIGITAL LED")
    cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")

    if st.button("ล้างการเลือกทั้งหมด"):
        st.session_state.selected_slides = set()
        st.session_state.selected_drive_images = set()
        st.rerun()

    st.divider()
    st.header("2. โหลดรูปจาก Google Drive")
    with st.expander("ℹ️ วิธีใช้ Google Drive", expanded=False):
        st.markdown("""
1. สร้าง API Key ที่ [Google Cloud Console](https://console.cloud.google.com/)
2. เปิดใช้ **Google Drive API**
3. โฟลเดอร์ต้องตั้งเป็น **"Anyone with the link"**
""")
    gdrive_api_key = st.text_input("Google Drive API Key", type="password", placeholder="AIza...")
    gdrive_folder_url = st.text_input("URL โฟลเดอร์ Drive", placeholder="https://drive.google.com/drive/folders/...")

    if st.button("📂 โหลดรูปจาก Drive"):
        if not gdrive_api_key or not gdrive_folder_url:
            st.error("กรุณากรอก API Key และ URL โฟลเดอร์")
        else:
            folder_id = extract_folder_id(gdrive_folder_url)
            if not folder_id:
                st.error("ไม่พบ Folder ID ในลิงก์นี้")
            else:
                with st.spinner("กำลังโหลดรายการรูป..."):
                    try:
                        files = list_drive_images(gdrive_api_key, folder_id)
                        st.session_state.drive_files = [
                            {**f, "api_key": gdrive_api_key} for f in files
                        ]
                        st.session_state.selected_drive_images = set()
                        st.success(f"พบรูป {len(files)} ไฟล์")
                    except requests.HTTPError as e:
                        st.error(f"Drive API Error: {e.response.status_code} — ตรวจสอบ API Key และสิทธิ์โฟลเดอร์")
                    except Exception as e:
                        st.error(f"เกิดข้อผิดพลาด: {e}")

# --- 4. แสดงสไลด์จากไฟล์ Master ---
if uploaded_file:
    master_obj, slides_data = process_file_optimized(uploaded_file.getvalue(), uploaded_file.name)

    st.header("2. เลือกหน้าจากไฟล์ Master")
    cols = st.columns(2)
    for idx, s in enumerate(slides_data):
        with cols[idx % 2]:
            with st.container(border=True):
                if s["preview"]:
                    st.image(s["preview"], use_container_width=True)
                is_checked = idx in st.session_state.selected_slides
                if st.checkbox(f"เลือก {s['display']}", key=f"cb_{idx}", value=is_checked):
                    st.session_state.selected_slides.add(idx)
                else:
                    st.session_state.selected_slides.discard(idx)
else:
    st.info("กรุณาอัปโหลดไฟล์ Master ที่แถบด้านซ้ายเพื่อเริ่มงานครับ")

# --- 5. แสดงรูปจาก Google Drive ---
if st.session_state.drive_files:
    st.header("3. เลือกรูปจาก Google Drive")
    drive_cols = st.columns(3)
    for idx, f in enumerate(st.session_state.drive_files):
        with drive_cols[idx % 3]:
            with st.container(border=True):
                with st.spinner(f"โหลด {f['name']}..."):
                    try:
                        img_bytes = download_drive_image(f["id"], f["api_key"])
                        st.image(img_bytes, use_container_width=True)
                    except Exception:
                        st.warning("โหลดรูปไม่ได้")
                        img_bytes = None
                is_checked = f["id"] in st.session_state.selected_drive_images
                if st.checkbox(f["name"], key=f"drive_cb_{idx}", value=is_checked):
                    st.session_state.selected_drive_images.add(f["id"])
                else:
                    st.session_state.selected_drive_images.discard(f["id"])

# --- 6. ขั้นตอนการสร้างไฟล์ผลลัพธ์ ---
has_selection = (
    (uploaded_file and st.session_state.selected_slides) or
    st.session_state.selected_drive_images
)

if has_selection:
    st.divider()
    if st.button("🏗️ เริ่มสร้างไฟล์ Appendix (16:9)", type="primary"):
        selected_list = sorted(list(st.session_state.selected_slides)) if uploaded_file else []
        selected_drive = list(st.session_state.selected_drive_images)
        total = len(selected_list) + len(selected_drive)

        progress_bar = st.progress(0, text="กำลังเริ่มสร้างไฟล์...")

        new_prs = Presentation()
        new_prs.slide_width = Inches(13.333)
        new_prs.slide_height = Inches(7.5)

        # หน้าปก
        cover_slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
        title_box = cover_slide.shapes.add_textbox(0, new_prs.slide_height / 3, new_prs.slide_width, Inches(1.5))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = cover_title
        p.alignment, p.font.size, p.font.bold = PP_ALIGN.CENTER, Pt(44), True

        done = 0

        # สไลด์จากไฟล์ Master
        if uploaded_file:
            slides_data_ref = process_file_optimized(uploaded_file.getvalue(), uploaded_file.name)[1]
            master_obj_ref = process_file_optimized(uploaded_file.getvalue(), uploaded_file.name)[0]
            for sid in selected_list:
                done += 1
                progress_bar.progress(done / total, text=f"ประมวลผลสไลด์ {done}/{total}")
                target_slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
                item = slides_data_ref[sid]
                if item["type"] == "pptx":
                    src_slide = master_obj_ref.slides[item["id"]]
                    for shp in src_slide.shapes:
                        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            add_full_image_16_9(target_slide, io.BytesIO(shp.image.blob), new_prs)
                        elif shp.has_text_frame:
                            copy_text_with_format(shp, target_slide)
                else:
                    page = master_obj_ref.load_page(item["id"])
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    add_full_image_16_9(target_slide, io.BytesIO(pix.tobytes("png")), new_prs)

        # รูปจาก Google Drive
        drive_file_map = {f["id"]: f for f in st.session_state.drive_files}
        for file_id in selected_drive:
            done += 1
            f = drive_file_map.get(file_id, {})
            progress_bar.progress(done / total, text=f"เพิ่มรูปจาก Drive: {f.get('name', file_id)} ({done}/{total})")
            try:
                img_bytes = download_drive_image(file_id, f["api_key"])
                target_slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
                add_full_image_16_9(target_slide, io.BytesIO(img_bytes), new_prs)
            except Exception as e:
                st.warning(f"ข้ามรูป {f.get('name', file_id)}: {e}")

        progress_bar.empty()
        output = io.BytesIO()
        new_prs.save(output)
        st.success("✅ สร้างไฟล์สัดส่วน 16:9 สำเร็จ!")
        st.download_button("📥 ดาวน์โหลดไฟล์ Appendix.pptx", output.getvalue(), "Appendix_16_9.pptx")
