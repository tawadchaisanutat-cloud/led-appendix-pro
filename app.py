import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
from PIL import Image
import fitz  # PyMuPDF

st.set_page_config(page_title="Digital LED Appendix Pro", layout="wide")

def copy_formatting(source_shape, target_shape):
    target_shape.text_frame.word_wrap = source_shape.text_frame.word_wrap
    for i, src_para in enumerate(source_shape.text_frame.paragraphs):
        target_para = target_shape.text_frame.paragraphs[0] if i == 0 else target_shape.text_frame.add_paragraph()
        target_para.alignment = src_para.alignment
        for src_run in src_para.runs:
            target_run = target_para.add_run()
            target_run.text = src_run.text
            if src_run.font.size: target_run.font.size = src_run.font.size
            if src_run.font.bold: target_run.font.bold = src_run.font.bold
            if src_run.font.name: target_run.font.name = src_run.font.name

def create_cover(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(0, prs.slide_height/2 - Inches(1), prs.slide_width, Inches(1.5))
    p = tb.text_frame.paragraphs[0]
    p.text = title
    p.alignment = PP_ALIGN.CENTER
    p.font.size, p.font.bold = Pt(44), True
    stb = slide.shapes.add_textbox(0, prs.slide_height/2 + Inches(0.5), prs.slide_width, Inches(1))
    sp = stb.text_frame.paragraphs[0]
    sp.text = subtitle
    sp.alignment = PP_ALIGN.CENTER
    sp.font.size = Pt(24)

st.title("🚀 Digital LED Appendix (Supports PDF & PPTX)")

with st.sidebar:
    st.header("1. Setup")
    uploaded_file = st.file_uploader("เลือกไฟล์ Master (.pptx หรือ .pdf)", type=["pptx", "pdf"])
    cover_title = st.text_input("หัวข้อหน้าปก", "PROPOSAL FOR DIGITAL LED")
    cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: [Customer Name]")

if uploaded_file:
    file_type = uploaded_file.name.split('.')[-1].lower()
    slides_data = []

    if file_type == "pptx":
        prs_master = Presentation(uploaded_file)
        for i, slide in enumerate(prs_master.slides):
            texts = [shape.text for shape in slide.shapes if hasattr(shape, "text")]
            preview_img = None
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    preview_img = io.BytesIO(shape.image.blob)
                    break
            slides_data.append({"id": i, "display": f"Slide {i+1}: {' '.join(texts)[:50]}", "preview": preview_img, "type": "pptx"})
    
    elif file_type == "pdf":
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        for i in range(len(doc)):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # Zoom 2x for clarity
            img_data = pix.tobytes("png")
            slides_data.append({"id": i, "display": f"PDF Page {i+1}", "preview": io.BytesIO(img_data), "type": "pdf"})

    st.header("2. Select Content")
    search_query = st.text_input("🔍 Search...")
    filtered = [s for s in slides_data if search_query.lower() in s["display"].lower()]
    
    selected_ids = []
    cols = st.columns(3)
    for idx, s in enumerate(filtered):
        with cols[idx % 3]:
            st.image(s["preview"], use_container_width=True)
            if st.checkbox(f"Select {s['display']}", key=f"s_{s['id']}"):
                selected_ids.append(s["id"])

    if st.button("Build Appendix", type="primary"):
        # สร้างไฟล์ PPTX ใหม่ (ใช้สัดส่วนมาตรฐาน 16:9)
        out_prs = Presentation() 
        create_cover(out_prs, cover_title, cover_sub)

        for sid in sorted(selected_ids):
            target_slide = out_prs.slides.add_slide(out_prs.slide_layouts[6])
            source = slides_data[sid]
            
            if source["type"] == "pptx":
                src_slide = prs_master.slides[sid]
                for shape in src_slide.shapes:
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        target_slide.shapes.add_picture(io.BytesIO(shape.image.blob), shape.left, shape.top, shape.width, shape.height)
                    elif shape.has_text_frame:
                        new_txt = target_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
                        copy_formatting(shape, new_txt)
            
            elif source["type"] == "pdf":
                # สำหรับ PDF เราจะวางรูปภาพเต็มหน้าสไลด์
                target_slide.shapes.add_picture(source["preview"], 0, 0, width=out_prs.slide_width, height=out_prs.slide_height)

        buf = io.BytesIO()
        out_prs.save(buf)
        st.download_button("📥 Download Appendix", buf.getvalue(), "Appendix.pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
