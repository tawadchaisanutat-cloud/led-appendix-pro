import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import re
import requests
from PIL import Image

st.set_page_config(page_title="LED Appendix Pro", layout="wide")

# --- Session state ---
if 'selected_images' not in st.session_state:
    st.session_state.selected_images = set()
if 'drive_files' not in st.session_state:
    st.session_state.drive_files = []
if 'loaded' not in st.session_state:
    st.session_state.loaded = False

# --- ดึง credentials จาก Secrets หรือ input ---
secrets = st.secrets if hasattr(st, 'secrets') else {}
API_KEY = secrets.get("GDRIVE_API_KEY", "")
FOLDER_ID = secrets.get("GDRIVE_FOLDER_ID", "")

# --- Google Drive helpers ---
def extract_folder_id(url):
    for pattern in [r'folders/([a-zA-Z0-9_-]+)', r'id=([a-zA-Z0-9_-]+)']:
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return url.strip()  # ถ้า user วาง ID ตรงๆ

@st.cache_data(show_spinner=False)
def list_drive_images(api_key, folder_id):
    all_files = []
    page_token = None
    while True:
        params = {
            "q": f"'{folder_id}' in parents and mimeType contains 'image/' and trashed=false",
            "key": api_key,
            "fields": "nextPageToken,files(id,name,mimeType)",
            "pageSize": 200,
        }
        if page_token:
            params["pageToken"] = page_token
        r = requests.get("https://www.googleapis.com/drive/v3/files", params=params, timeout=15)
        r.raise_for_status()
        data = r.json()
        all_files.extend(data.get("files", []))
        page_token = data.get("nextPageToken")
        if not page_token:
            break
    return sorted(all_files, key=lambda x: x["name"])

@st.cache_data(show_spinner=False)
def download_drive_image(file_id, api_key):
    r = requests.get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",
        params={"alt": "media", "key": api_key},
        timeout=30,
    )
    r.raise_for_status()
    return r.content

def add_full_image_16_9(slide, image_bytes, prs):
    img = Image.open(io.BytesIO(image_bytes))
    img_w, img_h = img.size
    ratio = min(prs.slide_width / img_w, prs.slide_height / img_h)
    new_w = int(img_w * ratio)
    new_h = int(img_h * ratio)
    left = (prs.slide_width - new_w) // 2
    top = (prs.slide_height - new_h) // 2
    slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width=new_w, height=new_h)

# --- Sidebar ---
with st.sidebar:
    st.title("LED Appendix Pro")
    st.divider()

    # ถ้าไม่มี Secrets ให้กรอกเอง
    if not API_KEY or not FOLDER_ID:
        st.subheader("ตั้งค่าการเชื่อมต่อ")
        with st.expander("ℹ️ วิธีตั้งค่า Secrets (ถาวร)", expanded=False):
            st.markdown("""
ใน Streamlit Cloud → App Settings → Secrets:
```toml
GDRIVE_API_KEY = "AIza..."
GDRIVE_FOLDER_ID = "1abc..."
```
ตั้งค่าครั้งเดียว ไม่ต้องกรอกซ้ำ
""")
        input_api_key = st.text_input("Google Drive API Key", type="password", placeholder="AIza...")
        input_folder = st.text_input("URL หรือ Folder ID", placeholder="https://drive.google.com/drive/folders/...")
        use_api_key = input_api_key
        use_folder_id = extract_folder_id(input_folder) if input_folder else ""
    else:
        use_api_key = API_KEY
        use_folder_id = FOLDER_ID
        st.success("เชื่อมต่อ Google Drive แล้ว")

    st.divider()

    # โหลดรูป
    auto_load = API_KEY and FOLDER_ID and not st.session_state.loaded
    if auto_load or st.button("🔄 โหลด / รีเฟรชรูป", use_container_width=True):
        if use_api_key and use_folder_id:
            with st.spinner("กำลังโหลดรายการรูป..."):
                try:
                    files = list_drive_images(use_api_key, use_folder_id)
                    st.session_state.drive_files = [
                        {**f, "api_key": use_api_key} for f in files
                    ]
                    st.session_state.loaded = True
                    st.session_state.selected_images = set()
                    st.rerun()
                except requests.HTTPError as e:
                    st.error(f"Drive API Error {e.response.status_code}: ตรวจสอบ API Key และสิทธิ์โฟลเดอร์")
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")
        else:
            st.warning("กรุณากรอก API Key และ Folder URL")

    if st.session_state.drive_files:
        st.info(f"รูปทั้งหมด {len(st.session_state.drive_files)} ไฟล์")
        st.info(f"เลือกแล้ว {len(st.session_state.selected_images)} ไฟล์")

        if st.button("ล้างการเลือกทั้งหมด", use_container_width=True):
            st.session_state.selected_images = set()
            st.rerun()

    st.divider()
    st.subheader("ตั้งค่าหน้าปก")
    cover_title = st.text_input("หัวข้อ", "PROPOSAL FOR DIGITAL LED")
    cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")

# --- Main area ---
st.title("LED Appendix Pro")

if not st.session_state.drive_files:
    st.info("กำลังรอโหลดรูปจาก Google Drive...")
    st.stop()

# --- ค้นหา ---
col_search, col_count = st.columns([3, 1])
with col_search:
    search = st.text_input("🔍 ค้นหารหัสภาพ", placeholder="เช่น DGT251, CENTRAL, NORTH...")
with col_count:
    st.metric("เลือกแล้ว", f"{len(st.session_state.selected_images)} / {len(st.session_state.drive_files)}")

# กรองรายการตาม search
all_files = st.session_state.drive_files
if search:
    filtered = [f for f in all_files if search.upper() in f["name"].upper()]
else:
    filtered = all_files

if not filtered:
    st.warning(f"ไม่พบรูปที่มีรหัส '{search}'")
    st.stop()

st.caption(f"แสดง {len(filtered)} รูป")

# --- Grid รูปภาพ ---
cols = st.columns(3)
for idx, f in enumerate(filtered):
    with cols[idx % 3]:
        with st.container(border=True):
            # Preview
            try:
                img_bytes = download_drive_image(f["id"], f["api_key"])
                st.image(img_bytes, use_container_width=True)
            except Exception:
                st.warning("โหลดไม่ได้")

            # Checkbox
            file_id = f["id"]
            is_checked = file_id in st.session_state.selected_images
            label = f["name"].replace(".jpeg", "").replace(".jpg", "").replace(".png", "")
            if st.checkbox(label, key=f"cb_{file_id}", value=is_checked):
                st.session_state.selected_images.add(file_id)
            else:
                st.session_state.selected_images.discard(file_id)

# --- Build ---
if st.session_state.selected_images:
    st.divider()
    if st.button(f"🏗️ สร้าง Appendix จาก {len(st.session_state.selected_images)} รูป (16:9)", type="primary", use_container_width=True):
        file_map = {f["id"]: f for f in all_files}
        selected_list = [file_map[fid] for fid in st.session_state.selected_images if fid in file_map]
        selected_list.sort(key=lambda x: x["name"])
        total = len(selected_list)

        progress = st.progress(0, text="กำลังเริ่มสร้างไฟล์...")

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        # หน้าปก
        cover = prs.slides.add_slide(prs.slide_layouts[6])
        tb = cover.shapes.add_textbox(0, prs.slide_height / 3, prs.slide_width, Inches(2))
        p = tb.text_frame.paragraphs[0]
        p.text = cover_title
        p.alignment, p.font.size, p.font.bold = PP_ALIGN.CENTER, Pt(44), True
        p2 = tb.text_frame.add_paragraph()
        p2.text = cover_sub
        p2.alignment, p2.font.size = PP_ALIGN.CENTER, Pt(24)

        # เพิ่มรูป
        errors = []
        for i, f in enumerate(selected_list):
            progress.progress((i + 1) / total, text=f"{f['name']} ({i+1}/{total})")
            try:
                img_bytes = download_drive_image(f["id"], f["api_key"])
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                add_full_image_16_9(slide, img_bytes, prs)
            except Exception as e:
                errors.append(f["name"])

        progress.empty()

        if errors:
            st.warning(f"ข้าม {len(errors)} รูปที่โหลดไม่ได้: {', '.join(errors)}")

        output = io.BytesIO()
        prs.save(output)
        st.success(f"✅ สร้างสำเร็จ! {total - len(errors)} สไลด์")
        st.download_button(
            "📥 ดาวน์โหลด Appendix.pptx",
            output.getvalue(),
            "Appendix_16_9.pptx",
            use_container_width=True,
        )
