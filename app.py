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

# --- 1. CONFIG ---
st.set_page_config(page_title="LED Appendix Pro", layout="wide", page_icon="🚀")

# ดึงค่าจาก Secrets ทันที
secrets = st.secrets if hasattr(st, 'secrets') else {}
API_KEY = secrets.get("GDRIVE_API_KEY", "")
FOLDER_ID = secrets.get("GDRIVE_FOLDER_ID", "")

# เตรียม Session State
if 'selected_images' not in st.session_state:
    st.session_state.selected_images = set()
if 'image_order' not in st.session_state:
    st.session_state.image_order = []

# --- 2. CORE FUNCTIONS (ดึงข้อมูลทันที) ---
@st.cache_data(show_spinner="🚀 กำลังดึงข้อมูลจาก Google Drive...", ttl=3600)
def fetch_all_drive_data(api_key, folder_id):
    """ฟังก์ชันที่จะถูกรันทันทีที่เปิดแอปเพื่อโหลดข้อมูล"""
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
    except:
        return []

@st.cache_data(show_spinner=False)
def download_file_content(file_id, api_key):
    r = requests.get(f"https://www.googleapis.com/drive/v3/files/{file_id}",
                     params={"alt": "media", "key": api_key}, timeout=30)
    return r.content

def add_img_to_slide(slide, image_bytes, prs):
    try:
        img = Image.open(io.BytesIO(image_bytes))
        w, h = img.size
        ratio = min(prs.slide_width / w, prs.slide_height / h)
        nw, nh = int(w * ratio), int(h * ratio)
        left, top = (prs.slide_width - nw) // 2, (prs.slide_height - nh) // 2
        slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width=nw, height=nh)
    except: pass

# --- 3. ⚡️ บังคับโหลดข้อมูล (ทำงานทันที) ---
# บรรทัดนี้ทำให้แอป 'ตื่นมาพร้อมข้อมูล' โดยไม่ต้องกดปุ่ม
all_files = fetch_all_drive_data(API_KEY, FOLDER_ID)

# --- 4. UI: MAIN SCREEN ---
st.title("🚀 LED Appendix Pro")

if not all_files:
    st.error("❌ ไม่สามารถดึงข้อมูลได้: โปรดเช็ก Secrets (API_KEY/FOLDER_ID) หรือสิทธิ์ของโฟลเดอร์")
    if st.button("ลองดึงข้อมูลใหม่อีกครั้ง"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

# ตั้งค่าหน้าปก
with st.expander("📝 ตั้งค่าหน้าปกและปกหลัง", expanded=False):
    c1, c2 = st.columns(2)
    c_title = c1.text_input("หัวข้อโปรเจกต์", "PROPOSAL FOR DIGITAL LED")
    c_sub = c2.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")

# ค้นหา CoverA/B
cover_a = next((f for f in all_files if "COVERA" in f['name'].upper()), None)
cover_b = next((f for f in all_files if "COVERB" in f['name'].upper()), None)

tab1, tab2 = st.tabs(["📷 เลือกรูปภาพ", f"🔀 จัดเรียง ({len(st.session_state.selected_images)})"])

# --- TAB 1: เลือกรูป ---
with tab1:
    search = st.text_input("🔍 ค้นหารหัสภาพ", placeholder="พิมพ์รหัสภาพ รูปจะขึ้นทันที...").upper()
    
    # Filter รูปภาพ
    display_files = [f for f in all_files if (search and search in f["name"].upper()) or (f["id"] in st.session_state.selected_images)]
    
    if not display_files:
        st.info("💡 พิมพ์รหัสภาพเพื่อค้นหา (รูปที่เลือกไว้จะค้างที่หน้าจอนี้เสมอ)")
    else:
        cols = st.columns(4)
        for idx, f in enumerate(display_files):
            with cols[idx % 4]:
                with st.container(border=True):
                    # Thumbnail จาก Google
                    t_url = f.get("thumbnailLink", "").replace("=s220", "=s400")
                    st.image(t_url, use_container_width=True)
                    
                    fid, fname = f["id"], f["name"]
                    is_sel = fid in st.session_state.selected_images
                    
                    if st.checkbox(fname[:15], key=f"c_{fid}", value=is_sel):
                        if fid not in st.session_state.selected_images:
                            st.session_state.selected_images.add(fid)
                            st.session_state.image_order.append(fid)
                    else:
                        if fid in st.session_state.selected_images:
                            st.session_state.selected_images.discard(fid)
                            st.session_state.image_order = [x for x in st.session_state.image_order if x != fid]

# --- TAB 2: จัดเรียง & EXPORT (แก้ไขจุด ValueError) ---
with tab2:
    # กรองเฉพาะ ID ที่ยังถูกเลือกอยู่
    current_order = [fid for fid in st.session_state.image_order if fid in st.session_state.selected_images]
    
    if not current_order:
        st.info("กรุณาเลือกรูปจากแท็บแรกก่อน")
    else:
        st.subheader("🖱️ คลิกลากเรียงลำดับ (บนสุด = สไลด์แรก)")
        id_to_name = {f['id']: f['name'] for f in all_files}
        
        # --- วิธีแก้ ValueError: บังคับให้เป็น List ของ String เท่านั้น ---
        sort_input = [f"{id_to_name.get(fid, 'Unknown')} | {fid}" for fid in current_order]
        
        # เรียกใช้ Sortables
        sorted_output = sort_items(sort_input, direction="vertical", key="final_order_fix")
        
        # ตัดเอา ID กลับมา
        new_id_order = [item.split(" | ")[-1] for item in sorted_output]
        st.session_state.image_order = new_id_order

        st.divider()
        if st.button("🏗️ สร้างไฟล์ PowerPoint", type="primary", use_container_width=True):
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            
            # 1. ปกหน้า
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            if cover_a:
                add_img_to_slide(slide, download_file_content(cover_a['id'], API_KEY), prs)
            tb = slide.shapes.add_textbox(0, prs.slide_height/2.5, prs.slide_width, Inches(2))
            p = tb.text_frame.paragraphs[0]
            p.text, p.alignment, p.font.size, p.font.bold = c_title, PP_ALIGN.CENTER, Pt(48), True
            if cover_a: p.font.color.rgb = RGBColor(255, 255, 255)
            
            # 2. เนื้อหา
            for fid in new_id_order:
                if (cover_a and fid == cover_a['id']) or (cover_b and fid == cover_b['id']): continue
                add_img_to_slide(prs.slides.add_slide(prs.slide_layouts[6]), download_file_content(fid, API_KEY), prs)
            
            # 3. ปกหลัง
            if cover_b:
                add_img_to_slide(prs.slides.add_slide(prs.slide_layouts[6]), download_file_content(cover_b['id'], API_KEY), prs)
            
            out = io.BytesIO()
            prs.save(out)
            st.success("🎉 สร้างไฟล์สำเร็จ!")
            st.download_button("📥 ดาวน์โหลด Appendix", out.getvalue(), "Appendix.pptx", use_container_width=True)

# --- 5. SIDEBAR ---
with st.sidebar:
    st.write(f"✅ โหลดรูปจาก Drive สำเร็จ: {len(all_files)} รูป")
    if st.button("🔄 ล้าง Cache และโหลดใหม่"):
        st.cache_data.clear()
        st.rerun()
