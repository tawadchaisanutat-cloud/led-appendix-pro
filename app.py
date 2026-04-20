import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re
import requests
from PIL import Image
from streamlit_sortables import sort_items 

# --- 1. SETTINGS & CONFIG ---
st.set_page_config(page_title="LED Appendix Pro", layout="wide", page_icon="🚀")

# Initialize Session State
if 'selected_images' not in st.session_state: st.session_state.selected_images = set()
if 'image_order' not in st.session_state: st.session_state.image_order = []
if 'drive_files' not in st.session_state: st.session_state.drive_files = []
if 'loaded' not in st.session_state: st.session_state.loaded = False

# ดึงค่าจาก Secrets
secrets = st.secrets if hasattr(st, 'secrets') else {}
API_KEY = secrets.get("GDRIVE_API_KEY", "")
FOLDER_ID = secrets.get("GDRIVE_FOLDER_ID", "")

# --- 2. HELPERS ---
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
    r.raise_for_status()
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

# --- 3. AUTO-LOAD LOGIC (จุดที่ทำให้ไม่ต้องกดปุ่ม) ---
if not st.session_state.loaded and API_KEY and FOLDER_ID:
    try:
        # สั่งโหลดทันทีถ้ามี Secrets
        files = list_drive_images(API_KEY, FOLDER_ID)
        st.session_state.drive_files = [{**f, "api_key": API_KEY} for f in files]
        st.session_state.loaded = True
        # ไม่ต้องสั่ง rerun เพื่อความเสถียร
    except Exception as e:
        st.error(f"การโหลดอัตโนมัติล้มเหลว: {e}")

# --- 4. UI: HEADER & COVER ---
st.title("🚀 LED Appendix Pro")

with st.expander("📝 ตั้งค่าข้อมูลหน้าปกและปกหลัง", expanded=True):
    col_c1, col_c2 = st.columns(2)
    with col_c1: cover_title = st.text_input("หัวข้อโปรเจกต์", "PROPOSAL FOR DIGITAL LED")
    with col_c2: cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")

# --- 5. SIDEBAR (คงไว้เผื่อโหลดใหม่) ---
with st.sidebar:
    st.subheader("⚙️ Settings")
    if not API_KEY or not FOLDER_ID:
        st.warning("กรุณาตั้งค่า Secrets ใน Streamlit Cloud")
    else:
        st.success("✅ เชื่อมต่ออัตโนมัติสำเร็จ")
    
    if st.button("🔄 รีเฟรชรูปใหม่", use_container_width=True):
        st.cache_data.clear()
        files = list_drive_images(API_KEY, FOLDER_ID)
        st.session_state.drive_files = [{**f, "api_key": API_KEY} for f in files]
        st.session_state.loaded = True
        st.rerun()

# --- 6. MAIN CONTENT ---
all_files = st.session_state.drive_files
if not all_files:
    st.info("กำลังดึงข้อมูลจาก Google Drive...")
    st.stop()

# หา Cover อัตโนมัติ
cover_a = next((f for f in all_files if "COVERA" in f['name'].upper()), None)
cover_b = next((f for f in all_files if "COVERB" in f['name'].upper()), None)

tab1, tab2 = st.tabs(["📷 เลือกรูปภาพ", f"🔀 จัดเรียง & Export ({len(st.session_state.selected_images)})"])

with tab1:
    search = st.text_input("🔍 ค้นหารหัสภาพ (เช่น DGT...)", placeholder="พิมพ์เพื่อค้นหา...").upper()
    display_files = [f for f in all_files if (search and search in f["name"].upper()) or (f["id"] in st.session_state.selected_images)]
    
    if not display_files:
        st.warning("ไม่พบรูป")
    else:
        cols = st.columns(4)
        for idx, f in enumerate(display_files):
            with cols[idx % 4]:
                with st.container(border=True):
                    t_url = f.get("thumbnailLink", "").replace("=s220", "=s400")
                    if t_url: st.image(t_url, use_container_width=True)
                    fid, fname = f["id"], f["name"]
                    is_sel = fid in st.session_state.selected_images
                    if st.checkbox(fname[:15], key=f"chk_{fid}", value=is_sel):
                        if fid not in st.session_state.selected_images:
                            st.session_state.selected_images.add(fid)
                            st.session_state.image_order.append(fid)
                    else:
                        if fid in st.session_state.selected_images:
                            st.session_state.selected_images.discard(fid)
                            st.session_state.image_order = [x for x in st.session_state.image_order if x != fid]

with tab2:
    current_order = [fid for fid in st.session_state.image_order if fid in st.session_state.selected_images]
    if not current_order:
        st.info("เลือกรูปภาพก่อนครับ")
    else:
        st.subheader("🖱️ คลิกลากเพื่อเรียงลำดับ")
        id_to_name = {f['id']: f['name'] for f in all_files}
        # แปลงเป็น String เพื่อกัน Error
        sort_input = [f"{id_to_name.get(fid, 'Unknown')} | {fid}" for fid in current_order]
        sorted_output = sort_items(sort_input, direction="vertical", key="final_sort")
        new_order = [item.split(" | ")[-1] for item in sorted_output]
        st.session_state.image_order = new_order

        if st.button("🏗️ สร้างไฟล์ PowerPoint", type="primary", use_container_width=True):
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            # --- สร้างหน้าปก ---
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            if cover_a:
                add_full_image_16_9(slide, download_drive_image(cover_a['id'], API_KEY), prs)
            tb = slide.shapes.add_textbox(0, prs.slide_height/2.5, prs.slide_width, Inches(2))
            p = tb.text_frame.paragraphs[0]
            p.text, p.alignment, p.font.size, p.font.bold = cover_title, PP_ALIGN.CENTER, Pt(48), True
            if cover_a: p.font.color.rgb = RGBColor(255, 255, 255)
            # --- ใส่รูปเนื้อหา ---
            for fid in new_order:
                if (cover_a and fid == cover_a['id']) or (cover_b and fid == cover_b['id']): continue
                add_full_image_16_9(prs.slides.add_slide(prs.slide_layouts[6]), download_drive_image(fid, API_KEY), prs)
            # --- ใส่ปกหลัง ---
            if cover_b:
                add_full_image_16_9(prs.slides.add_slide(prs.slide_layouts[6]), download_drive_image(cover_b['id'], API_KEY), prs)
            
            ppt_out = io.BytesIO()
            prs.save(ppt_out)
            st.download_button("📥 ดาวน์โหลด Appendix", ppt_out.getvalue(), "Appendix.pptx", use_container_width=True)
