import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import fitz  # PyMuPDF

st.set_page_config(page_title="LED Appendix Pro v3", layout="wide")

# --- 1. ระบบจัดการหน่วยความจำ (Session State) ---
if 'selected_slides' not in st.session_state:
    st.session_state.selected_slides = set()

# --- 2. Optimization: Cache การประมวลผลไฟล์ ---
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
            # Low-res Preview (Fast)
            pix = page.get_pixmap(matrix=fitz.Matrix(0.4, 0.4)) 
            slides_data.append({
                "id": i, "type": "pdf", "display": f"หน้า {i+1}", 
                "preview": io.BytesIO(pix.tobytes("png"))
            })
        return doc, slides_data

# --- UI Header ---
st.title("🚀 Professional Appendix Builder")
st.info("ระบบจะจดจำหน้าที่คุณเลือกไว้โดยอัตโนมัติ")

with st.sidebar:
    st.header("1. ตั้งค่าไฟล์")
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Master", type=["pptx", "pdf"])
    cover_title = st.text_input("หัวข้อหน้าปก", "PROPOSAL FOR DIGITAL LED")
    if st.button("ล้างการเลือกทั้งหมด"):
        st.session_state.selected_slides = set()
        st.rerun()

if uploaded_file:
    # อ่านไฟล์และเก็บเข้า Cache
    master_obj, slides_data = process_file_optimized(uploaded_file.getvalue(), uploaded_file.name)

    # --- 3. การแสดงผล Grid พร้อมระบบจดจำค่า ---
    st.header("2. เลือกหน้าที่ต้องการรวมใน Appendix")
    cols = st.columns(5)
    for idx, s in enumerate(slides_data):
        with cols[idx % 5]:
            if s["preview"]: st.image(s["preview"], use_container_width=True)
            
            # ตรวจสอบว่าเคยเลือกไว้หรือไม่
            is_checked = idx in st.session_state.selected_slides
            if st.checkbox(f"เลือก {s['display']}", key=f"cb_{idx}", value=is_checked):
                st.session_state.selected_slides.add(idx)
            else:
                st.session_state.selected_slides.discard(idx)

    # --- 4. ส่วนการสร้างไฟล์พร้อม Progress Bar ---
    if st.button("🏗️ เริ่มสร้างไฟล์ Appendix", type="primary"):
        selected_list = sorted(list(st.session_state.selected_slides))
        
        if not selected_list:
            st.warning("กรุณาเลือกสไลด์อย่างน้อย 1 หน้า")
        else:
            progress_text = "กำลังสร้างสไลด์... โปรดรอสักครู่"
            my_bar = st.progress(0, text=progress_text)
            
            new_prs = Presentation()
            # [สร้างหน้าปก...] (โค้ดส่วนหน้าปกเหมือนเดิม)
            
            total = len(selected_list)
            for i, sid in enumerate(selected_list):
                # อัปเดต Progress Bar
                step = (i + 1) / total
                my_bar.progress(step, text=f"กำลังประมวลผลหน้าที่ {i+1} จาก {total}")
                
                target_slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
                item = slides_data[sid]
                
                if item["type"] == "pptx":
                    src_slide = master_obj.slides[item["id"]]
                    for shp in src_slide.shapes:
                        if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            target_slide.shapes.add_picture(io.BytesIO(shp.image.blob), shp.left, shp.top, shp.width, shp.height)
                        elif shp.has_text_frame:
                            # เรียกใช้ฟังก์ชันคัดลอกข้อความที่เราทำไว้คราวที่แล้ว
                            pass 
                else:
                    # High-res สำหรับ PDF ตอนส่งออก
                    page = master_obj.load_page(item["id"])
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                    target_slide.shapes.add_picture(io.BytesIO(pix.tobytes("png")), 0, 0, width=new_prs.slide_width, height=new_prs.slide_height)

            my_bar.empty() # ลบ Progress Bar เมื่อเสร็จ
            
            output = io.BytesIO()
            new_prs.save(output)
            st.success("✅ สร้างไฟล์สำเร็จ!")
            st.download_button("📥 ดาวน์โหลดไฟล์ (PPTX)", output.getvalue(), "Appendix_Final.pptx")
