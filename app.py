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

# เริ่มต้นตัวแปรใน Session State
if 'selected_images' not in st.session_state:
    st.session_state.selected_images = set()
if 'image_order' not in st.session_state:
    st.session_state.image_order = []
if 'drive_files' not in st.session_state:
    st.session_state.drive_files = []
if 'loaded' not in st.session_state:
    st.session_state.loaded = False

secrets = st.secrets if hasattr(st, 'secrets') else {}
API_KEY = secrets.get("GDRIVE_API_KEY", "")
FOLDER_ID = secrets.get("GDRIVE_FOLDER_ID", "")

# --- 2. HELPERS ---
def extract_folder_id(url):
    for pattern in [r'folders/([a-zA-Z0-9_-]+)', r'id=([a-zA-Z0-9_-]+)']:
        m = re.search(pattern, url)
        if m: return m.group(1)
    return url.strip()

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
    except Exception:
        pass

# --- 3. UI: HEADER ---
st.title("🚀 LED Appendix Pro")

# ส่วนตั้งค่าหน้าปก (ย้ายขึ้นมาด้านบน)
with st.expander("📝 ตั้งค่าข้อมูลหน้าปกและปกหลัง", expanded=True):
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        cover_title = st.text_input("หัวข้อโปรเจกต์ (Main Title)", "PROPOSAL FOR DIGITAL LED")
    with col_c2:
        cover_sub = st.text_input("ชื่อลูกค้า/ผู้รับ (Subtitle)", "Presented to: Valued Customer")
    st.caption("💡 ระบบจะค้นหาไฟล์ 'CoverA' สำหรับปกหน้า และ 'CoverB' สำหรับปกหลังให้อัตโนมัติ")

# --- 4. SIDEBAR ---
with st.sidebar:
    st.subheader("⚙️ การเชื่อมต่อ")
    if not API_KEY or not FOLDER_ID:
        u_api = st.text_input("API Key", type="password")
        u_fid = st.text_input("Folder ID / URL")
        use_api_key = u_api
        use_folder_id = extract_folder_id(u_fid) if u_fid else ""
    else:
        use_api_key, use_folder_id = API_KEY, FOLDER_ID
        st.success("✅ API Connected")

    if st.button("🔄 โหลด/รีเฟรชรูป", use_container_width=True):
        if use_api_key and use_folder_id:
            st.cache_data.clear()
            files = list_drive_images(use_api_key, use_folder_id)
            st.session_state.drive_files = [{**f, "api_key": use_api_key} for f in files]
            st.session_state.loaded = True
            st.rerun()

# --- 5. LOGIC: PRE-PROCESSING ---
all_files = st.session_state.drive_files
# หา Cover อัตโนมัติ (ไม่สนตัวเล็กตัวใหญ่)
cover_a = next((f for f in all_files if "COVERA" in f['name'].upper()), None)
cover_b = next((f for f in all_files if "COVERB" in f['name'].upper()), None)

if not st.session_state.loaded and use_api_key and use_folder_id:
    files = list_drive_images(use_api_key, use_folder_id)
    st.session_state.drive_files = [{**f, "api_key": use_api_key} for f in files]
    st.session_state.loaded = True
    st.rerun()

if not all_files:
    st.info("กรุณากดปุ่ม 'โหลด/รีเฟรชรูป' ที่แถบด้านข้างเพื่อเริ่มดึงข้อมูล")
    st.stop()

tab1, tab2 = st.tabs(["📷 เลือกรูปภาพ", f"🔀 จัดเรียง & Export ({len(st.session_state.selected_images)})"])

# --- TAB 1: เลือกรูป ---
with tab1:
    search = st.text_input("🔍 ค้นหารหัสภาพ", placeholder="พิมพ์ชื่อภาพที่ต้องการ...").upper()
    
    # กรองเฉพาะรูปที่ค้นหาเจอ หรือรูปที่เคยเลือกไว้แล้ว
    display_files = [
        f for f in all_files 
        if (search and search in f["name"].upper()) or (f["id"] in st.session_state.selected_images)
    ]

    if not display_files:
        st.warning("ไม่พบรูปที่ค้นหา หรือยังไม่ได้เลือกรูปใดๆ")
    else:
        cols = st.columns(4)
        for idx, f in enumerate(display_files):
            with cols[idx % 4]:
                with st.container(border=True):
                    # ใช้ Thumbnail ของ Google (เร็วมาก)
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

# --- TAB 2: จัดเรียงลำดับ (ลากวาง) ---
with tab2:
    # กรองเอาเฉพาะ ID ที่ยังมีตัวตนอยู่ในชุดที่เลือก
    current_order = [fid for fid in st.session_state.image_order if fid in st.session_state.selected_images]
    
    if not current_order:
        st.info("กรุณาเลือกรูปภาพจากแท็บแรกก่อน")
    else:
        st.subheader("🖱️ คลิกลากเพื่อเรียงลำดับภาพ")
        id_to_name = {f['id']: f['name'] for f in all_files}
        
        # แก้ไข ValueError: แปลงข้อมูลให้เป็น List ของ String เท่านั้น
        sort_input = [f"{id_to_name.get(fid, 'Unknown')} | {fid}" for fid in current_order]
        
        # เรียกใช้งาน Sortables (ส่งเฉพาะ List[str])
        sorted_output = sort_items(sort_input, direction="vertical", key="sort_list_final")
        
        # ตัดเอา ID กลับมา
        new_order = [item.split(" | ")[-1] for item in sorted_output]
        st.session_state.image_order = new_order

        st.divider()
        
        if st.button("🏗️ สร้างไฟล์ PowerPoint (.pptx)", type="primary", use_container_width=True):
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            
            p_bar = st.progress(0, text="กำลังสร้างไฟล์...")
            total_steps = len(new_order) + 2 # ปกหน้า + รูป + ปกหลัง
            step = 0

            # 1. ปกหน้า
            step += 1
            p_bar.progress(step/total_steps, text="กำลังสร้างหน้าปก...")
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            if cover_a:
                img_data = download_drive_image(cover_a['id'], use_api_key)
                add_full_image_16_9(slide, img_data, prs)
            
            # ใส่ข้อความหน้าปก
            tb = slide.shapes.add_textbox(0, prs.slide_height/2.5, prs.slide_width, Inches(2))
            p = tb.text_frame.paragraphs[0]
            p.text, p.alignment, p.font.size, p.font.bold = cover_title, PP_ALIGN.CENTER, Pt(48), True
            if cover_a: p.font.color.rgb = RGBColor(255, 255, 255)
            
            p2 = tb.text_frame.add_paragraph()
            p2.text, p2.alignment, p2.font.size = cover_sub, PP_ALIGN.CENTER, Pt(26)
            if cover_a: p2.font.color.rgb = RGBColor(255, 255, 255)

            # 2. รูปภาพเนื้อหา
            for fid in new_order:
                # ข้ามรูปที่เป็นปกเพื่อไม่ให้ซ้ำ
                if (cover_a and fid == cover_a['id']) or (cover_b and fid == cover_b['id']):
                    continue
                step += 1
                p_bar.progress(step/total_steps, text=f"กำลังใส่รูป: {id_to_name.get(fid)}")
                img_data = download_drive_image(fid, use_api_key)
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                add_full_image_16_9(slide, img_data, prs)

            # 3. ปกหลัง
            if cover_b:
                step += 1
                p_bar.progress(1.0, text="กำลังใส่ปกหลัง...")
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                img_data = download_drive_image(cover_b['id'], use_api_key)
                add_full_image_16_9(slide, img_data, prs)

            # Save
            ppt_out = io.BytesIO()
            prs.save(ppt_out)
            st.success("🎉 สร้างไฟล์สำเร็จ!")
            st.download_button("📥 ดาวน์โหลด Appendix", ppt_out.getvalue(), "Appendix_16_9.pptx", use_container_width=True)
