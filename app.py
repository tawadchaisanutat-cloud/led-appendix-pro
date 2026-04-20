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
for key, default in [
    ('selected_images', set()), # เก็บ Image IDs ที่ถูกเลือก
    ('image_order', []),        # เก็บ Image IDs แบบเรียงลำดับ
    ('drive_files', []),        # รายการไฟล์ทั้งหมดจาก Drive
    ('loaded', False),          # สถานะการโหลดครั้งแรก
]:
    if key not in st.session_state:
        st.session_state[key] = default

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
    except Exception as e:
        st.error(f"Error adding image: {e}")

# --- 3. UI: HEADER & COVER SETTINGS ---
st.title("🚀 LED Appendix Pro")

# ส่วนตั้งชื่อหน้าปก (ย้ายขึ้นมาด้านบนสุด)
with st.container(border=True):
    st.subheader("📝 ตั้งค่าหน้าปก")
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        cover_title = st.text_input("หัวข้อโปรเจกต์ (Main Title)", "PROPOSAL FOR DIGITAL LED")
    with col_c2:
        cover_sub = st.text_input("ชื่อลูกค้า/ผู้รับ (Subtitle)", "Presented to: Valued Customer")
    st.caption("💡 ระบบจะดึงรูปชื่อ 'CoverA' เป็นปกหน้า และ 'CoverB' เป็นปกหลังโดยอัตโนมัติ")

# --- 4. SIDEBAR ---
with st.sidebar:
    st.header("⚙️ การเชื่อมต่อ")
    if not API_KEY or not FOLDER_ID:
        input_api = st.text_input("API Key", type="password")
        input_fid = st.text_input("Folder URL/ID")
        use_api_key = input_api
        use_folder_id = extract_folder_id(input_fid) if input_fid else ""
    else:
        use_api_key, use_folder_id = API_KEY, FOLDER_ID
        st.success("✅ Connected via Secrets")

    if st.button("🔄 โหลด/รีเฟรชรูป", use_container_width=True):
        if use_api_key and use_folder_id:
            st.cache_data.clear()
            files = list_drive_images(use_api_key, use_folder_id)
            st.session_state.drive_files = [{**f, "api_key": use_api_key} for f in files]
            st.session_state.loaded = True
            st.rerun()
        else:
            st.error("กรุณากรอก API Key และ Folder ID")

# --- 5. LOGIC: PRE-PROCESSING ---
all_files = st.session_state.drive_files
# หา Cover อัตโนมัติ
cover_a = next((f for f in all_files if "COVERA" in f['name'].upper()), None)
cover_b = next((f for f in all_files if "COVERB" in f['name'].upper()), None)

# --- 6. MAIN TABS ---
if not st.session_state.loaded and use_api_key and use_folder_id:
    # Auto-load ครั้งแรกถ้ามีข้อมูลครบ
    files = list_drive_images(use_api_key, use_folder_id)
    st.session_state.drive_files = [{**f, "api_key": use_api_key} for f in files]
    st.session_state.loaded = True
    st.rerun()

if not all_files:
    st.info("กรุณากด 'โหลด/รีเฟรชรูป' เพื่อเริ่มดึงข้อมูลจาก Google Drive")
    st.stop()

tab1, tab2 = st.tabs(["📷 เลือกรูปภาพ", f"🔀 จัดเรียง & Export ({len(st.session_state.selected_images)})"])

# --- TAB 1: เลือกรูป (เฉพาะที่ค้นหา + ที่เลือกแล้ว) ---
with tab1:
    search = st.text_input("🔍 ค้นหารหัสภาพ", placeholder="พิมพ์เพื่อค้นหา...").upper()
    
    # Filter: โชว์เฉพาะที่ค้นหาเจอ หรือ ที่ถูกเลือกไว้แล้วเท่านั้น
    display_files = [
        f for f in all_files 
        if (search and search in f["name"].upper()) or (f["id"] in st.session_state.selected_images)
    ]

    if not display_files:
        st.warning("ไม่พบรูปที่ค้นหา หรือยังไม่ได้เลือกรูปใดๆ")
    else:
        st.caption(f"แสดงทั้งหมด {len(display_files)} รูป (ค้นหาพบ / เลือกไว้)")
        cols = st.columns(4) # ปรับเป็น 2 สำหรับ Mobile ถ้าต้องการ
        for idx, f in enumerate(display_files):
            with cols[idx % 4]:
                with st.container(border=True):
                    # ใช้ ThumbnailLink จาก Google (เร็วสุด)
                    thumb = f.get("thumbnailLink", "").replace("=s220", "=s400")
                    st.image(thumb, use_container_width=True)
                    
                    fid = f["id"]
                    fname = f["name"]
                    is_checked = fid in st.session_state.selected_images
                    
                    if st.checkbox(fname[:20], key=f"chk_{fid}", value=is_checked):
                        if fid not in st.session_state.selected_images:
                            st.session_state.selected_images.add(fid)
                            st.session_state.image_order.append(fid)
                    else:
                        if fid in st.session_state.selected_images:
                            st.session_state.selected_images.discard(fid)
                            st.session_state.image_order = [x for x in st.session_state.image_order if x != fid]

# --- TAB 2: จัดเรียงลำดับ (ลากวาง) & EXPORT ---
with tab2:
    # กรองเอาเฉพาะ ID ที่ยังคงถูกเลือกอยู่
    current_order = [fid for fid in st.session_state.image_order if fid in st.session_state.selected_images]
    
    if not current_order:
        st.info("กรุณาไปเลือกรูปในแท็บ 'เลือกรูปภาพ' ก่อนครับ")
    else:
        st.subheader("🖱️ คลิกลากเพื่อเรียงลำดับ (บนสุดจะอยู่สไลด์แรก)")
        
        # เตรียมข้อมูล String สำหรับ sort_items (Fix ValueError)
        id_to_name = {f['id']: f['name'] for f in all_files}
        # สร้าง String format: "ชื่อไฟล์ | ID"
        sort_input = [f"{id_to_name.get(fid, 'Unknown')} | {fid}" for fid in current_order]
        
        # ตัวลากวาง
        sorted_output = sort_items(sort_input, direction="vertical", key="main_sort_list")
        
        # แกะ ID กลับออกมาหลังลากเสร็จ
        new_order = [item.split(" | ")[-1] for item in sorted_output]
        st.session_state.image_order = new_order

        st.divider()

        # --- EXPORT PROCESS ---
        if st.button("🏗️ สร้างไฟล์ PowerPoint (.pptx)", type="primary", use_container_width=True):
            prs = Presentation()
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)
            
            progress_text = "กำลังเริ่มสร้างไฟล์..."
            progress_bar = st.progress(0, text=progress_text)
            
            # คำนวณจำนวนขั้นตอน (CoverA + Images + CoverB)
            total_slides = len(new_order) + (1 if cover_a else 1) + (1 if cover_b else 0)
            current_step = 0

            # 1. ปกหน้า
            current_step += 1
            progress_bar.progress(current_step/total_slides, text="กำลังสร้างหน้าปก...")
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank layout
            
            if cover_a:
                img_bytes = download_drive_image(cover_a['id'], use_api_key)
                add_full_image_16_9(slide, img_bytes, prs)
            
            # ใส่ Text ทับหน้าปก
            tb = slide.shapes.add_textbox(0, prs.slide_height/2.5, prs.slide_width, Inches(3))
            p = tb.text_frame.paragraphs[0]
            p.text = cover_title
            p.alignment, p.font.size, p.font.bold = PP_ALIGN.CENTER, Pt(50), True
            if cover_a: p.font.color.rgb = RGBColor(255, 255, 255) # สีขาวถ้ามีรูปพื้นหลัง

            p2 = tb.text_frame.add_paragraph()
            p2.text = cover_sub
            p2.alignment, p2.font.size = PP_ALIGN.CENTER, Pt(28)
            if cover_a: p2.font.color.rgb = RGBColor(255, 255, 255)

            # 2. เนื้อหา (รูปที่เลือก)
            for fid in new_order:
                # ข้ามรูป CoverA/B ถ้าบังเอิญไปติ๊กเลือกมา เพื่อไม่ให้ซ้ำ
                if (cover_a and fid == cover_a['id']) or (cover_b and fid == cover_b['id']):
                    continue
                
                current_step += 1
                fname = id_to_name.get(fid, "Image")
                progress_bar.progress(current_step/total_slides, text=f"กำลังใส่รูป: {fname}")
                
                img_bytes = download_drive_image(fid, use_api_key)
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                add_full_image_16_9(slide, img_bytes, prs)

            # 3. ปกหลัง (Auto)
            if cover_b:
                current_step += 1
                progress_bar.progress(current_step/total_slides, text="กำลังใส่ปกหลัง...")
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                img_bytes = download_drive_image(cover_b['id'], use_api_key)
                add_full_image_16_9(slide, img_bytes, prs)

            # Save and Download
            ppt_out = io.BytesIO()
            prs.save(ppt_out)
            progress_bar.empty()
            st.success(f"🎉 สร้างสำเร็จทั้งหมด {current_step} สไลด์!")
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ Appendix.pptx",
                data=ppt_out.getvalue(),
                file_name="LED_Appendix_16_9.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
