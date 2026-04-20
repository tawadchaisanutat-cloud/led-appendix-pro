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

# ดึงค่าจาก Secrets ทันที
secrets = st.secrets if hasattr(st, 'secrets') else {}
API_KEY = secrets.get("GDRIVE_API_KEY", "")
FOLDER_ID = secrets.get("GDRIVE_FOLDER_ID", "")

# เตรียม Session State สำหรับการเลือกและจัดลำดับเท่านั้น
if 'selected_images' not in st.session_state:
    st.session_state.selected_images = set()
if 'image_order' not in st.session_state:
    st.session_state.image_order = []

# --- 2. HELPERS (มีการ Cache ข้อมูลไว้ในเครื่อง) ---
@st.cache_data(show_spinner="กำลังเชื่อมต่อ Google Drive...", ttl=3600)
def fetch_all_files(api_key, folder_id):
    """ฟังก์ชันหลักที่ดึงข้อมูลทันทีและเก็บไว้ใน Cache"""
    if not api_key or not folder_id:
        return []
    all_files, page_token = [], None
    try:
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
    except Exception as e:
        st.error(f"Error fetching data: {e}")
        return []

@st.cache_data(show_spinner=False)
def download_image(file_id, api_key):
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

# --- 3. ⚡️ EXECUTION (ดึงข้อมูลทันทีทุกครั้งที่รันแอป) ---
# เรียกใช้ฟังก์ชัน fetch โดยตรง (ถ้ามี Cache มันจะดึงมาทันทีในเสี้ยววินาที)
all_files = fetch_all_files(API_KEY, FOLDER_ID)

# --- 4. UI: MAIN SCREEN ---
st.title("🚀 LED Appendix Pro")

if not API_KEY or not FOLDER_ID:
    st.error("❌ ไม่พบข้อมูลการเชื่อมต่อใน Secrets (GDRIVE_API_KEY / GDRIVE_FOLDER_ID)")
    st.stop()

if not all_files:
    st.warning("⚠️ ไม่พบไฟล์ภาพใน Folder ที่ระบุ หรือ API Key ไม่ถูกต้อง")
    if st.button("ลองโหลดใหม่อีกครั้ง"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

# ส่วนตั้งค่าหน้าปก
with st.expander("📝 ตั้งค่าข้อมูลหน้าปกและปกหลัง", expanded=False):
    col_c1, col_c2 = st.columns(2)
    with col_c1: cover_title = st.text_input("หัวข้อโปรเจกต์", "PROPOSAL FOR DIGITAL LED")
    with col_c2: cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")

# ค้นหาไฟล์ Cover อัตโนมัติ
cover_a = next((f for f in all_files if "COVERA" in f['name'].upper()), None)
cover_b = next((f for f in all_files if "COVERB" in f['name'].upper()), None)

tab1, tab2 = st.tabs(["📷 เลือกรูปภาพ", f"🔀 จัดเรียง & Export ({len(st.session_state.selected_images)})"])

# --- TAB 1: เลือกรูป ---
with tab1:
    search = st.text_input("🔍 ค้นหารหัสภาพ", placeholder="พิมพ์รหัสภาพที่ต้องการที่นี่...").upper()
    
    # กรองรูปที่จะแสดง
    display_files = [f for f in all_files if (search and search in f["name"].upper()) or (f["id"] in st.session_state.selected_images)]
    
    if not display_files:
        st.info("💡 พิมพ์รหัสภาพในช่องด้านบนเพื่อค้นหา (รูปที่เลือกไว้จะแสดงที่นี่เสมอ)")
    else:
        st.caption(f"พบข้อมูล {len(display_files)} รายการ")
        cols = st.columns(4)
        for idx, f in enumerate(display_files):
            with cols[idx % 4]:
                with st.container(border=True):
                    # Thumbnail Link (Google จัดการให้ เร็วมาก)
                    t_url = f.get("thumbnailLink", "").replace("=s220", "=s400")
                    st.image(t_url, use_container_width=True)
                    
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

# --- TAB 2: จัดเรียง & EXPORT ---
with tab2:
    # กรองเฉพาะ ID ที่ยังถูกเลือกอยู่
    current_order = [fid for fid in st.session_state.image_order if fid in st.session_state.selected_images]
    
    if not current_order:
        st.info("กรุณาเลือกรูปภาพจากแท็บแรกก่อน")
    else:
        st.subheader("🖱️ คลิกลากเพื่อเรียงลำดับสไลด์")
        id_to_name = {f['id']: f['name'] for f in all_files}
        
        # จัดการข้อมูลสำหรับ Sortables (Fixed ValueError)
        sort_input = [f"{id_to_name.get(fid, 'Unknown')} | {fid}" for fid in current_order]
        sorted_output = sort_items(sort_input, direction="vertical", key="sort_final_v5")
        
        # อัปเดตลำดับ
        new_order = [item.split(" | ")[-1] for item in sorted_output]
        st.session_state.image_order = new_order

        st.divider()
        if st.button("🏗️ สร้างไฟล์ PowerPoint (.pptx)", type="primary", use_container_width=True):
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            # ปกหน้า
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            if cover_a: add_full_image_16_9(slide, download_image(cover_a['id'], API_KEY), prs)
            tb = slide.shapes.add_textbox(0, prs.slide_height/2.5, prs.slide_width, Inches(2))
            p = tb.text_frame.paragraphs[0]
            p.text, p.alignment, p.font.size, p.font.bold = cover_title, PP_ALIGN.CENTER, Pt(48), True
            if cover_a: p.font.color.rgb = RGBColor(255, 255, 255)
            # เนื้อหา
            for fid in new_order:
                if (cover_a and fid == cover_a['id']) or (cover_b and fid == cover_b['id']): continue
                add_full_image_16_9(prs.slides.add_slide(prs.slide_layouts[6]), download_image(fid, API_KEY), prs)
            # ปกหลัง
            if cover_b: add_full_image_16_9(prs.slides.add_slide(prs.slide_layouts[6]), download_image(cover_b['id'], API_KEY), prs)
            
            ppt_out = io.BytesIO()
            prs.save(ppt_out)
            st.download_button("📥 ดาวน์โหลด Appendix", ppt_out.getvalue(), "Appendix.pptx", use_container_width=True)

# --- 5. SIDEBAR ---
with st.sidebar:
    st.subheader("⚙️ Settings")
    st.success(f"✅ โหลดไฟล์จาก Drive สำเร็จ ({len(all_files)} รูป)")
    if st.button("🔄 ล้าง Cache และโหลดใหม่"):
        st.cache_data.clear()
        st.rerun()
