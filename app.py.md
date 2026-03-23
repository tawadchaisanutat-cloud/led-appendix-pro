import streamlit as st  
from pptx import Presentation  
from pptx.enum.shapes import MSO_SHAPE_TYPE  
from pptx.util import Inches, Pt  
from pptx.enum.text import PP_ALIGN  
import io  
from PIL import Image  
  
# --- Page Config ---  
st.set_page_config(page_title="Digital LED Appendix Pro", layout="wide")  
  
def copy_formatting(source_shape, target_shape):  
    """คัดลอกรูปแบบข้อความระดับ Paragraph และ Run"""  
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
            try:  
                if src_run.font.color.rgb: target_run.font.color.rgb = src_run.font.color.rgb  
            except: pass  
  
def create_cover(prs, title, subtitle):  
    slide = prs.slides.add_slide(prs.slide_layouts[6])  
    w, h = prs.slide_width, prs.slide_height  
    # Title  
    tb = slide.shapes.add_textbox(0, h/2 - Inches(1), w, Inches(1.5))  
    p = tb.text_frame.paragraphs[0]  
    p.text = title  
    p.alignment = PP_ALIGN.CENTER  
    p.font.size, p.font.bold = Pt(44), True  
    # Subtitle  
    stb = slide.shapes.add_textbox(0, h/2 + Inches(0.5), w, Inches(1))  
    sp = stb.text_frame.paragraphs[0]  
    sp.text = subtitle  
    sp.alignment = PP_ALIGN.CENTER  
    sp.font.size = Pt(24)  
  
# --- UI Layout ---  
st.title("🚀 Digital LED Appendix Builder")  
st.markdown("ระบบสร้างเอกสาร Appendix สำหรับ iPad และ Web Browser")  
  
# Sidebar สำหรับการตั้งค่า  
with st.sidebar:  
    st.header("1. ตั้งค่าไฟล์และหน้าปก")  
    uploaded_file = st.file_uploader("เลือกไฟล์ Master (.pptx)", type="pptx")  
      
    cover_title = st.text_input("หัวข้อหน้าปก", "PROPOSAL FOR DIGITAL LED")  
    cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: [Customer Name]")  
      
    st.divider()  
    st.info("คัดลอกรูปภาพและข้อความจากสไลด์ต้นฉบับมาสร้างไฟล์ใหม่")  
  
if uploaded_file:  
    # โหลดไฟล์ PPTX เข้า Memory  
    prs = Presentation(uploaded_file)  
    slides_data = []  
  
    for i, slide in enumerate(prs.slides):  
        texts = [shape.text for shape in slide.shapes if hasattr(shape, "text")]  
        full_text = " ".join(texts).replace('\n', ' ')  
          
        # หา Preview Image  
        preview_img = None  
        for shape in slide.shapes:  
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:  
                preview_img = io.BytesIO(shape.image.blob)  
                break  
          
        slides_data.append({  
            "id": i,  
            "display": f"Slide {i+1}: {full_text[:50]}...",  
            "preview": preview_img  
        })  
  
    # --- ส่วนการค้นหาและเลือกสไลด์ ---  
    st.header("2. ค้นหาและเลือกสไลด์")  
    search_query = st.text_input("🔍 ค้นหาข้อความในสไลด์...")  
      
    filtered_slides = [s for s in slides_data if search_query.lower() in s["display"].lower()]  
      
    # แสดงสไลด์แบบ Grid ให้เลือก (เหมาะกับหน้าจอสัมผัส iPad)  
    selected_ids = []  
    cols = st.columns(3) # แสดง 3 คอลัมน์  
      
    for idx, s in enumerate(filtered_slides):  
        with cols[idx % 3]:  
            if s["preview"]:  
                st.image(s["preview"], use_container_width=True)  
            else:  
                st.warning("ไม่มีรูปพรีวิว")  
              
            if st.checkbox(f"เลือก {s['display'][:20]}", key=f"slide_{s['id']}"):  
                selected_ids.append(s["id"])  
  
    st.divider()  
    st.write(f"เลือกไว้ทั้งหมด: **{len(selected_ids)}** สไลด์")  
  
    # --- ส่วนการ Generate ไฟล์ ---  
    if st.button("🏗️ เริ่มสร้างไฟล์ Appendix", type="primary", use_container_width=True):  
        if not selected_ids:  
            st.error("กรุณาเลือกสไลด์อย่างน้อย 1 รายการ")  
        else:  
            # สร้าง PPTX ใหม่ในหน่วยความจำ  
            output_prs = Presentation(uploaded_file) # ใช้ต้นฉบับเป็นฐานเพื่อรักษา Theme  
            # ลบสไลด์เก่า  
            xml_slides = output_prs.slides._sldIdLst  
            for _ in range(len(xml_slides)):  
                xml_slides.remove(xml_slides[0])  
  
            # สร้างหน้าปก  
            create_cover(output_prs, cover_title, cover_sub)  
  
            # เพิ่มสไลด์ที่เลือก  
            for sid in sorted(selected_ids):  
                source_slide = prs.slides[sid]  
                new_slide = output_prs.slides.add_slide(output_prs.slide_layouts[6])  
                for shape in source_slide.shapes:  
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:  
                        new_slide.shapes.add_picture(io.BytesIO(shape.image.blob),   
                                                   shape.left, shape.top, shape.width, shape.height)  
                    elif shape.has_text_frame:  
                        new_txt = new_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)  
                        copy_formatting(shape, new_txt)  
  
            # บันทึกลงหน่วยความจำเพื่อให้ดาวน์โหลดได้  
            binary_output = io.BytesIO()  
            output_prs.save(binary_output)  
            binary_output.seek(0)  
  
            st.success("สร้างไฟล์สำเร็จ!")  
            st.download_button(  
                label="📥 ดาวน์โหลดไฟล์ Appendix.pptx",  
                data=binary_output,  
                file_name="Appendix_Export.pptx",  
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",  
                use_container_width=True  
            )  
else:  
    st.info("กรุณาอัปโหลดไฟล์ Master PPTX เพื่อเริ่มต้น")  
