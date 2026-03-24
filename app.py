import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
from PIL import Image
import fitz  # Library สำหรับ PDF

st.set_page_config(page_title="LED Appendix Pro (PDF & PPTX)", layout="wide")

# ฟังก์ชันคัดลอกรูปแบบข้อความ
def copy_formatting(src_shape, tgt_shape):
    tgt_shape.text_frame.word_wrap = src_shape.text_frame.word_wrap
    for i, src_para in enumerate(src_shape.text_frame.paragraphs):
        tgt_para = tgt_shape.text_frame.paragraphs[0] if i == 0 else tgt_shape.text_frame.add_paragraph()
        tgt_para.alignment = src_para.alignment
        for src_run in src_para.runs:
            tgt_run = tgt_para.add_run()
            tgt_run.text = src_run.text
            if src_run.font.size: tgt_run.font.size = src_run.font.size
            if src_run.font.bold: tgt_run.font.bold = src_run.font.bold

st.title("🚀 Digital LED Appendix Builder")

with st.sidebar:
    st.header("1. ตั้งค่าไฟล์")
    uploaded_file = st.file_uploader("เลือกไฟล์ Master (.pptx หรือ .pdf)", type=["pptx", "pdf"])
    cover_title = st.text_input("หัวข้อหน้าปก", "PROPOSAL FOR DIGITAL LED")
    cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: [Customer Name]")

if uploaded_file:
    f_ext = uploaded_file.name.split('.')[-1].lower()
    slides_data = []

    if f_ext == "pptx":
        master_prs = Presentation(uploaded_file)
        for i, slide in enumerate(master_prs.slides):
            preview = None
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    preview = io.BytesIO(shape.image.blob)
                    break
            slides_data.append({"id": i, "type": "pptx", "display": f"Slide {i+1}", "preview": preview})
    
    elif f_ext == "pdf":
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        for i in range(len(doc)):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_data = pix.tobytes("png")
            slides_data.append({"id": i, "type": "pdf", "display": f"PDF หน้า {i+1}", "preview": io.BytesIO(img_data)})

    st.header("2. เลือกหน้าที่ต้องการ")
    selected_ids = []
    cols = st.columns(3)
    for idx, s in enumerate(slides_data):
        with cols[idx % 3]:
            if s["preview"]: st.image(s["preview"], use_container_width=True)
            if st.checkbox(f"เลือก {s['display']}", key=f"s_{idx}"):
                selected_ids.append(idx)

    if st.button("🏗️ สร้างไฟล์ Appendix", type="primary"):
        new_prs = Presentation()
        # หน้าปก (สร้างแบบง่าย)
        slide = new_prs.slides.add_slide(new_prs.slide_layouts[6])
        box = slide.shapes.add_textbox(0, new_prs.slide_height/2 - Inches(1), new_prs.slide_width, Inches(2))
        p = box.text_frame.paragraphs[0]
        p.text = cover_title
        p.alignment = PP_ALIGN.CENTER
        p.font.size, p.font.bold = Pt(44), True

        for sid in selected_ids:
            target = new_prs.slides.add_slide(new_prs.slide_layouts[6])
            item = slides_data[sid]
            if item["type"] == "pptx":
                # คัดลอกรูปจากสไลด์
                src_slide = master_prs.slides[item["id"]]
                for shp in src_slide.shapes:
                    if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        target.shapes.add_picture(io.BytesIO(shp.image.blob), shp.left, shp.top, shp.width, shp.height)
            else:
                # วาง PDF เป็นรูปเต็มหน้า
                target.shapes.add_picture(item["preview"], 0, 0, width=new_prs.slide_width, height=new_prs.slide_height)

        output = io.BytesIO()
        new_prs.save(output)
        st.download_button("📥 ดาวน์โหลดผลลัพธ์ (PPTX)", output.getvalue(), "Appendix.pptx")
