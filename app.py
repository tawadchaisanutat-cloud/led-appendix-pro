import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import re
import requests
from PIL import Image
# ต้องติดตั้ง: pip install streamlit-sortables
from streamlit_sortables import sort_items 

# --- Config ---
st.set_page_config(page_title="LED Appendix Pro", layout="wide")

# --- Session state ---
for key, default in [
    ('selected_images', set()), # เก็บ IDs
    ('image_order', []),        # เก็บ IDs แบบเรียงลำดับ
    ('drive_files', []),
    ('loaded', False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

secrets = st.secrets if hasattr(st, 'secrets') else {}
API_KEY = secrets.get("GDRIVE_API_KEY", "")
FOLDER_ID = secrets.get("GDRIVE_FOLDER_ID", "")

# --- Helpers ---
@st.cache_data(show_spinner=False, ttl=3600)
def list_drive_images(api_key, folder_id):
    all_files, page_token = [], None
    while True:
        params = {
            "q": f"'{folder_id}' in parents and mimeType contains 'image/' and trashed=false",
            "key": api_key,
            "fields": "nextPageToken,files(id,name,thumbnailLink)",
            "pageSize": 1000,
        }
        r = requests.get("https://www.googleapis.com/drive/v3/files", params=params, timeout=15)
        r.raise_for_status()
        data = r.json()
        all_files.extend(data.get("files", []))
        page_token = data.get("nextPageToken")
        if not page_token: break
    return sorted(all_files, key=lambda x: x["name"])

@st.cache_data(show_spinner=False)
def download_drive_image(file_id, api_key):
    r = requests.get(f"https://www.googleapis.com/drive/v3/files/{file_id}",
                     params={"alt": "media", "key": api_key}, timeout=30)
    return r.content

def add_full_image_16_9(slide, image_bytes, prs):
    try:
        img = Image.open(io.BytesIO(image_bytes))
        img_w, img_h = img.size
        ratio = min(prs.slide_width / img_w, prs.slide_height / img_h)
        new_w, new_h = int(img_w * ratio), int(img_h * ratio)
        left, top = (prs.slide_width - new_w) // 2, (prs.slide_height - new_h) // 2
        slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width=new_w, height=new_h)
    except: pass

# --- UI: Top Header & Cover Settings ---
st.title("🚀 LED Appendix Pro")

# 1. ย้ายการตั้งค่าหน้าปกขึ้นมาด้านบน
with st.expander("📝 ตั้งค่าข้อมูลหน้าปกและปกหลัง", expanded=True):
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        cover_title = st.text_input("หัวข้อโปรเจกต์ (Main Title)", "PROPOSAL FOR DIGITAL LED")
    with col_c2:
        cover_sub = st.text_input("ชื่อลูกค้า/ผู้รับ (Subtitle)", "Presented to: Valued Customer")
    st.caption("💡 ระบบจะค้นหาไฟล์ชื่อ 'CoverA' สำหรับปกหน้า และ 'CoverB' สำหรับปกหลังให้อัตโนมัติ")

# --- Sidebar ---
with st.sidebar:
    st.subheader("Settings")
    if not API_KEY or not FOLDER_ID:
        use_api_key = st.text_input("API Key", type="password")
        use_folder_id = st.text_input("Folder ID")
    else:
        use_api_key, use_folder_id = API_KEY, FOLDER_ID
        st.success("API Connected")
    
    if st.button("🔄 โหลด/รีเฟรชรูป", use_container_width=True):
        st.cache_data.clear()
        files = list_drive_images(use_api_key, use_folder_id)
        st.session_state.drive_files = [{**f, "api_key": use_api_key} for f in files]
        st.session_state.loaded = True
        st.rerun()

# --- Logic: Auto Cover Detection ---
all_files = st.session_state.drive_files
cover_a = next((f for f in all_files if "CoverA" in f['name'].upper()), None)
cover_b = next((f for f in all_files if "CoverB" in f['name'].upper()), None)

# --- Tabs ---
tab1, tab2 = st.tabs(["📷 เลือกรูปภาพ", "📂 จัดเรียงลำดับ & Export"])

# --- TAB 1: เลือกรูป (เฉพาะที่ค้นหา + ที่เลือกแล้ว) ---
with tab1:
    search = st.text_input("🔍 ค้นหารหัสภาพ", placeholder="พิมพ์ชื่อภาพที่ต้องการ...").upper()
    
    # Filter: โชว์เฉพาะที่ตรงกับ Search OR โชว์เฉพาะที่เคยเลือกไว้แล้ว
    display_files = [
        f for f in all_files 
        if (search and search in f["name"].upper()) or (f["id"] in st.session_state.selected_images)
    ]

    if not display_files:
        st.info("กรุณาพิมพ์ค้นหารูปภาพ (เช่น DGT...) รูปที่เลือกไว้จะปรากฏที่นี่ด้วย")
    else:
        st.caption(f"แสดงทั้งหมด {len(display_files)} รูป (รวมที่เลือกไว้)")
        cols = st.columns(4)
        for idx, f in enumerate(display_files):
            with cols[idx % 4]:
                with st.container(border=True):
                    thumb = f.get("thumbnailLink", "").replace("=s220", "=s400")
                    st.image(thumb, use_container_width=True)
                    
                    is_sel = f["id"] in st.session_state.selected_images
                    if st.checkbox(f['name'][:15], key=f"sel_{f['id']}", value=is_sel):
                        if f["id"] not in st.session_state.selected_images:
                            st.session_state.selected_images.add(f["id"])
                            st.session_state.image_order.append(f["id"])
                    else:
                        if f["id"] in st.session_state.selected_images:
                            st.session_state.selected_images.discard(f["id"])
                            st.session_state.image_order = [x for x in st.session_state.image_order if x != f["id"]]

# --- TAB 2: จัดเรียงลำดับ (Drag & Drop) ---
with tab2:
    current_order = [fid for fid in st.session_state.image_order if fid in st.session_state.selected_images]
    
    if not current_order:
        st.warning("ยังไม่มีรูปที่ถูกเลือก กรุณากลับไปเลือกที่แท็บ 'เลือกรูปภาพ'")
    else:
        st.subheader("🖱️ คลิกลากเพื่อเรียงลำดับภาพ")
        
        # สร้าง List ของชื่อภาพเพื่อใช้ในการ Sort
        id_to_name = {f['id']: f['name'] for f in all_files}
        sort_data = [{"id": fid, "content": id_to_name.get(fid, fid)} for fid in current_order]
        
        # ใช้ streamlit-sortables
        sorted_data = sort_items(sort_data, direction="vertical", key="sortable_list")
        
        # อัปเดตลำดับกลับเข้า session_state
        new_order = [item["id"] for item in sorted_data]
        st.session_state.image_order = new_order

        st.divider()
        
        # --- Export Section ---
        if st.button("🏗️ สร้างไฟล์ PPTX ตอนนี้", type="primary", use_container_width=True):
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            
            progress_bar = st.progress(0)
            total_steps = len(new_order) + (1 if cover_a else 0) + (1 if cover_b else 0)
            step = 0

            # 1. หน้าปก (Auto CoverA หรือ Text)
            step += 1
            progress_bar.progress(step/total_steps, "กำลังสร้างหน้าปก...")
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            if cover_a:
                img_data = download_drive_image(cover_a['id'], use_api_key)
                add_full_image_16_9(slide, img_data, prs)
            
            # ใส่ข้อความทับหน้าปก
            tb = slide.shapes.add_textbox(0, prs.slide_height/3, prs.slide_width, Inches(2))
            p = tb.text_frame.paragraphs[0]
            p.text, p.alignment, p.font.size, p.font.bold = cover_title, PP_ALIGN.CENTER, Pt(44), True
            p.font.color.rgb = (255, 255, 255) if cover_a else (0,0,0)
            
            # 2. รูปภาพที่จัดเรียงไว้ (ข้าม CoverA/B ถ้าอยู่ใน List เพื่อไม่ให้ซ้ำ)
            for fid in new_order:
                if (cover_a and fid == cover_a['id']) or (cover_b and fid == cover_b['id']):
                    continue
                step += 1
                progress_bar.progress(step/total_steps, f"กำลังใส่รูป: {id_to_name[fid]}")
                img_data = download_drive_image(fid, use_api_key)
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                add_full_image_16_9(slide, img_data, prs)

            # 3. ปกหลัง (Auto CoverB)
            if cover_b:
                step += 1
                progress_bar.progress(step/total_steps, "กำลังสร้างปกหลัง...")
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                img_data = download_drive_image(cover_b['id'], use_api_key)
                add_full_image_16_9(slide, img_data, prs)

            # Save
            output = io.BytesIO()
            prs.save(output)
            st.success("🎉 สร้างไฟล์เสร็จเรียบร้อย!")
            st.download_button("📥 ดาวน์โหลด Appendix", output.getvalue(), "LED_Appendix.pptx", use_container_width=True)
