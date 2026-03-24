import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import fitz  # PyMuPDF
from PIL import Image

# --- ตั้งค่าหน้ากระดาษ ---
st.set_page_config(page_title="LED Appendix Pro v3", layout="wide")

# --- 1. ระบบจัดการสถานะ (Session State) ---
if 'selected_slides' not in st.session_state:
    st.session_state.selected_slides = set()

# --- 2. ฟังก์ชันประมวลผล (Optimization & Cache) ---
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
            # ปรับความชัด Preview เพิ่มขึ้น (0.8) เพื่อให้อ่านรหัสได้ชัดเจนขึ้น
            pix = page.get_pixmap(matrix=fitz.Matrix(0.8, 0.8)) 
            slides_data.append({
                "id": i, "type": "pdf", "display": f"หน้า {i+1}", 
                "preview": io.BytesIO(pix.tobytes("png"))
            })
        return doc, slides_data

def copy_text_with_format(src_shape, target_slide):
    """คัดลอกข้อความพร้อมรูปแบบจากต้นฉบับ"""
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

def add_full_image(slide, image_stream, prs):
    """วางรูปให้ใหญ่ที่สุดกึ่งกลางสไลด์"""
    img = Image.open(image_stream)
    img_w, img_h = img.size
    ratio = min(prs.slide_width / img_w, prs.slide_height / img_h)
    new_w, new_h = int(img_w * ratio), int(img_h * ratio)
    left = (prs.slide_width - new_w) // 2
    top = (prs.slide_height - new_h) // 2
    image_stream.seek(0)
    slide.shapes.add_picture(image_stream, left, top, width=new_w, height=new_h)

# --- 3. ส่วน UI (แก้ NameError โดยรับค่าก่อนเช็คเงื่อนไข) ---
st.title("🚀 LED Appendix Builder Pro")

with st.sidebar:
    st.header("1. ตั้งค่าไฟล์")
    # ประกาศตัวแปรรับค่าไฟล์ก่อนเสมอ
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Master (PPTX/PDF)", type=["pptx", "pdf"])
    cover_title = st.text_input("หัวข้อหน้าปก", "PROPOSAL FOR DIGITAL LED")
    cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")
    
    if st.button("ล้างการเลือกทั้งหมด"):
        st.session_state.selected_slides = set()
        st.rerun()

# ตรวจสอบว่ามีไฟล์ถูกอัปโหลดหรือยัง
if uploaded_file:
    # อ่านไฟล์และเก็บเข้า Cache
    master_obj, slides_data = process_file_optimized(uploaded_file.getvalue(), uploaded_file.name)

    st.header("2. เลือกหน้าที่ต้องการ (แสดงผลขนาดใหญ่เพื่อให้เห็นรหัสชัดเจน)")
    
    # ปรับเป็น 2 คอลัมน์เพื่อให้รูปภาพใหญ่ขึ้น
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

    # --- 4. การสร้างไฟล์ ---
    if st.button("🏗️ เริ่มสร้างไฟล์ Appendix", type="primary"):
        selected_list = sorted(list(st.session_state.selected_slides))
        if not selected_list:
            st.warning("กรุณาเลือกสไลด์อย่างน้อย 1 หน้า")
        else:
            progress_bar = st.progress(0, text="กำลังเริ่มสร้างไฟล์...")
            new_prs = Presentation()
            
            # สร้างหน้าปก
            cover_slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
            title_box = cover_slide.shapes.add_textbox(0, new_prs.slide_height/3, new_prs.slide_width, Inches(1))
            tf = title_box.text_frame
            tf.text = cover_title
            p = tf.paragraphs[0]
            p.alignment, p.font.size, p.font.bold = PP_ALIGN.CENTER, Pt(40), True

            # ใส่เนื้อหา
            total = len(selected_list)
            for i, sid in enumerate(selected_list):
                progress_bar.progress((i + 1) / total, text=f"กำลังประมวลผลหน้าที่ {i+1}/{total}")
                target_slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
                item = slides_data[sid]
                
                if item["type"] == "pptx":
                    src_slide = master_obj.slides[item["id"]]
                    for shp in src_slide.shapes:
                        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            add_full_image(target_slide, io.BytesIO(shp.image.blob), new_prs)
                        elif shp.has_text_frame:
                            copy_text_with_format(shp, target_slide)
                else:
                    # PDF High-res
                    page = master_obj.load_page(item["id"])
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    add_full_image(target_slide, io.BytesIO(pix.tobytes("png")), new_prs)

            progress_bar.empty()
            output = io.BytesIO()
            new_prs.save(output)
            st.success("✅ สร้างไฟล์สำเร็จ!")
            st.download_button("📥 ดาวน์โหลดไฟล์ (PPTX)", output.getvalue(), "Appendix_Final.pptx")
else:
    st.info("กรุณาอัปโหลดไฟล์ที่แถบด้านซ้ายเพื่อเริ่มใช้งานครับ")
