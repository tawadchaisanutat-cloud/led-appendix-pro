import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import fitz  # PyMuPDF
from PIL import Image

# --- การตั้งค่าหน้าเว็บ ---
st.set_page_config(page_title="LED Appendix Pro v3 (16:9)", layout="wide")

# --- 1. ระบบจัดการสถานะและ Cache ---
if 'selected_slides' not in st.session_state:
    st.session_state.selected_slides = set()

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
            # ปรับความชัด Preview (0.8) เพื่อให้อ่านรหัสได้ชัดเจน
            pix = page.get_pixmap(matrix=fitz.Matrix(0.8, 0.8)) 
            slides_data.append({
                "id": i, "type": "pdf", "display": f"หน้า {i+1}", 
                "preview": io.BytesIO(pix.tobytes("png"))
            })
        return doc, slides_data

# --- 2. ฟังก์ชันช่วยจัดการ Layout ---
def copy_text_with_format(src_shape, target_slide):
    """คัดลอกข้อความและรักษาฟอร์แมตพื้นฐาน"""
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
    """วางรูปภาพให้ขยายเต็มสัดส่วนสไลด์ (Fit to 16:9)"""
    img = Image.open(image_stream)
    img_w, img_h = img.size
    
    # คำนวณอัตราส่วนเพื่อให้ภาพใหญ่ที่สุดในพื้นที่ 16:9
    ratio = min(prs.slide_width / img_w, prs.slide_height / img_h)
    new_w, new_h = int(img_w * ratio), int(img_h * ratio)
    
    # จัดวางกึ่งกลาง
    left = (prs.slide_width - new_w) // 2
    top = (prs.slide_height - new_h) // 2
    
    image_stream.seek(0)
    slide.shapes.add_picture(image_stream, left, top, width=new_w, height=new_h)

# --- 3. ส่วน UI และ Logic หลัก ---
st.title("🚀 Professional LED Appendix Builder (16:9)")

with st.sidebar:
    st.header("1. ตั้งค่าไฟล์")
    # รับค่าไฟล์ก่อนเพื่อป้องกัน NameError
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Master (PPTX/PDF)", type=["pptx", "pdf"])
    cover_title = st.text_input("หัวข้อหน้าปก", "PROPOSAL FOR DIGITAL LED")
    cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")
    
    if st.button("ล้างการเลือกทั้งหมด"):
        st.session_state.selected_slides = set()
        st.rerun()

if uploaded_file:
    master_obj, slides_data = process_file_optimized(uploaded_file.getvalue(), uploaded_file.name)

    st.header("2. เลือกหน้าที่ต้องการ (มุมมองขยายเพื่อรหัสที่ชัดเจน)")
    
    # ใช้ 2 คอลัมน์เพื่อให้ภาพ Preview ใหญ่พอจะอ่านรหัสได้
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

    # --- 4. ขั้นตอนการสร้างไฟล์ผลลัพธ์ ---
    if st.button("🏗️ เริ่มสร้างไฟล์ Appendix (16:9)", type="primary"):
        selected_list = sorted(list(st.session_state.selected_slides))
        if not selected_list:
            st.warning("กรุณาเลือกสไลด์อย่างน้อย 1 หน้า")
        else:
            progress_bar = st.progress(0, text="กำลังเริ่มสร้างไฟล์...")
            
            # สร้าง Presentation และกำหนดขนาดเป็น 16:9 (Widescreen)
            new_prs = Presentation()
            new_prs.slide_width = Inches(13.333)
            new_prs.slide_height = Inches(7.5)
            
            # สร้างหน้าปก
            cover_slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
            title_box = cover_slide.shapes.add_textbox(0, new_prs.slide_height/3, new_prs.slide_width, Inches(1.5))
            tf = title_box.text_frame
            p = tf.paragraphs[0]
            p.text = cover_title
            p.alignment, p.font.size, p.font.bold = PP_ALIGN.CENTER, Pt(44), True
            
            # ใส่สไลด์ที่เลือก
            total = len(selected_list)
            for i, sid in enumerate(selected_list):
                progress_bar.progress((i + 1) / total, text=f"กำลังประมวลผลหน้า {i+1}/{total}")
                target_slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
                item = slides_data[sid]
                
                if item["type"] == "pptx":
                    src_slide = master_obj.slides[item["id"]]
                    for shp in src_slide.shapes:
                        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            add_full_image_16_9(target_slide, io.BytesIO(shp.image.blob), new_prs)
                        elif shp.has_text_frame:
                            copy_text_with_format(shp, target_slide)
                else:
                    # กรณี PDF: ดึงภาพความละเอียดสูงแบบ 16:9
                    page = master_obj.load_page(item["id"])
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    add_full_image_16_9(target_slide, io.BytesIO(pix.tobytes("png")), new_prs)

            progress_bar.empty()
            output = io.BytesIO()
            new_prs.save(output)
            st.success("✅ สร้างไฟล์สัดส่วน 16:9 สำเร็จ!")
            st.download_button("📥 ดาวน์โหลดไฟล์ Appendix.pptx", output.getvalue(), "Appendix_16_9.pptx")
else:
    st.info("กรุณาอัปโหลดไฟล์ Master ที่แถบด้านซ้ายเพื่อเริ่มงานครับ")
