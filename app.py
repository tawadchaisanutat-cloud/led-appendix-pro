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
for key, default in [
    ('selected_images', set()),
    ('image_order', []),
    ('drive_files', []),
    ('loaded', False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# --- Credentials ---
secrets = st.secrets if hasattr(st, 'secrets') else {}
API_KEY = secrets.get("GDRIVE_API_KEY", "")
FOLDER_ID = secrets.get("GDRIVE_FOLDER_ID", "")

# --- Google Drive helpers ---
def extract_folder_id(url):
    for pattern in [r'folders/([a-zA-Z0-9_-]+)', r'id=([a-zA-Z0-9_-]+)']:
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return url.strip()

@st.cache_data(show_spinner=False)
def list_drive_images(api_key, folder_id):
    all_files, page_token = [], None
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
    new_w, new_h = int(img_w * ratio), int(img_h * ratio)
    left = (prs.slide_width - new_w) // 2
    top = (prs.slide_height - new_h) // 2
    slide.shapes.add_picture(io.BytesIO(image_bytes), left, top, width=new_w, height=new_h)

# --- Order management ---
def sync_order():
    sel = st.session_state.selected_images
    st.session_state.image_order = [fid for fid in st.session_state.image_order if fid in sel]
    for fid in sel:
        if fid not in st.session_state.image_order:
            st.session_state.image_order.append(fid)

def move_up(idx):
    o = st.session_state.image_order
    if idx > 0:
        o[idx], o[idx - 1] = o[idx - 1], o[idx]

def move_down(idx):
    o = st.session_state.image_order
    if idx < len(o) - 1:
        o[idx], o[idx + 1] = o[idx + 1], o[idx]

def remove_from_order(fid):
    st.session_state.image_order = [x for x in st.session_state.image_order if x != fid]
    st.session_state.selected_images.discard(fid)

# --- Sidebar ---
with st.sidebar:
    st.title("LED Appendix Pro")
    st.divider()

    if not API_KEY or not FOLDER_ID:
        st.subheader("ตั้งค่าการเชื่อมต่อ")
        with st.expander("ℹ️ ตั้งค่า Secrets (ถาวร)", expanded=False):
            st.markdown("""
Streamlit Cloud → Settings → Secrets:
```toml
GDRIVE_API_KEY = "AIza..."
GDRIVE_FOLDER_ID = "1abc..."
```
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

    auto_load = API_KEY and FOLDER_ID and not st.session_state.loaded
    if auto_load or st.button("🔄 โหลด / รีเฟรชรูป", use_container_width=True):
        if use_api_key and use_folder_id:
            with st.spinner("กำลังโหลดรายการรูป..."):
                try:
                    files = list_drive_images(use_api_key, use_folder_id)
                    st.session_state.drive_files = [{**f, "api_key": use_api_key} for f in files]
                    st.session_state.loaded = True
                    st.session_state.selected_images = set()
                    st.session_state.image_order = []
                    st.rerun()
                except requests.HTTPError as e:
                    st.error(f"Drive API Error {e.response.status_code}")
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")
        else:
            st.warning("กรุณากรอก API Key และ Folder URL")

    if st.session_state.drive_files:
        st.info(f"รูปทั้งหมด {len(st.session_state.drive_files)} ไฟล์")
        selected_count = len(st.session_state.selected_images)
        if selected_count:
            st.success(f"เลือกแล้ว {selected_count} รูป")
        if st.button("ล้างการเลือกทั้งหมด", use_container_width=True):
            st.session_state.selected_images = set()
            st.session_state.image_order = []
            st.rerun()

    st.divider()
    st.subheader("ตั้งค่าหน้าปก")
    cover_title = st.text_input("หัวข้อ", "PROPOSAL FOR DIGITAL LED")
    cover_sub = st.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")

# --- Main ---
st.title("LED Appendix Pro")

if not st.session_state.drive_files:
    st.info("กำลังรอโหลดรูปจาก Google Drive...")
    st.stop()

n_selected = len(st.session_state.selected_images)
tab1, tab2 = st.tabs([
    "📷 เลือกรูป",
    f"🔀 จัดเรียง & Preview ({n_selected} รูป)",
])

# ===================== TAB 1: เลือกรูป =====================
with tab1:
    col_search, col_metric = st.columns([3, 1])
    with col_search:
        search = st.text_input("🔍 ค้นหารหัสภาพ", placeholder="เช่น DGT251, CENTRAL, NORTH...")
    with col_metric:
        st.metric("เลือกแล้ว", f"{n_selected} รูป")

    all_files = st.session_state.drive_files
    filtered = [f for f in all_files if search.upper() in f["name"].upper()] if search else all_files

    if not filtered:
        st.warning(f"ไม่พบรูปที่มีรหัส '{search}'")
    else:
        st.caption(f"แสดง {len(filtered)} รูป")
        cols = st.columns(3)
        for idx, f in enumerate(filtered):
            with cols[idx % 3]:
                with st.container(border=True):
                    try:
                        img_bytes = download_drive_image(f["id"], f["api_key"])
                        st.image(img_bytes, use_container_width=True)
                    except Exception:
                        st.warning("โหลดไม่ได้")
                    label = f["name"].rsplit(".", 1)[0]
                    fid = f["id"]
                    checked = fid in st.session_state.selected_images
                    if st.checkbox(label, key=f"cb_{fid}", value=checked):
                        st.session_state.selected_images.add(fid)
                    else:
                        st.session_state.selected_images.discard(fid)

    if n_selected:
        st.divider()
        st.info(f"เลือก {n_selected} รูปแล้ว → ไปที่แท็บ **จัดเรียง & Preview** เพื่อดูและเรียงลำดับก่อน Export")

# ===================== TAB 2: จัดเรียง & Preview =====================
with tab2:
    sync_order()
    order = st.session_state.image_order

    if not order:
        st.info("ยังไม่ได้เลือกรูป — กลับไปแท็บ **เลือกรูป** แล้วติ๊กรูปที่ต้องการครับ")
        st.stop()

    file_map = {f["id"]: f for f in st.session_state.drive_files}
    total = len(order)

    st.caption(f"รูปที่เลือก {total} รูป — กด ↑ ↓ เพื่อเรียงลำดับ, กด ✕ เพื่อลบออก")
    st.divider()

    for i, fid in enumerate(order):
        f = file_map.get(fid)
        if not f:
            continue

        col_num, col_img, col_name, col_up, col_dn, col_rm = st.columns([0.4, 1.5, 4, 0.5, 0.5, 0.5])

        with col_num:
            st.markdown(f"<div style='font-size:1.4rem;font-weight:bold;padding-top:24px;text-align:center'>{i+1}</div>", unsafe_allow_html=True)

        with col_img:
            try:
                img_bytes = download_drive_image(fid, f["api_key"])
                st.image(img_bytes, use_container_width=True)
            except Exception:
                st.write("❌")

        with col_name:
            st.markdown(f"<div style='padding-top:28px'>{f['name'].rsplit('.', 1)[0]}</div>", unsafe_allow_html=True)

        with col_up:
            st.button("↑", key=f"up_{fid}", on_click=move_up, args=(i,), disabled=(i == 0), use_container_width=True)

        with col_dn:
            st.button("↓", key=f"dn_{fid}", on_click=move_down, args=(i,), disabled=(i == total - 1), use_container_width=True)

        with col_rm:
            st.button("✕", key=f"rm_{fid}", on_click=remove_from_order, args=(fid,), use_container_width=True)

        st.divider()

    # --- Build ---
    st.subheader(f"Export {total} สไลด์ เป็น PPTX 16:9")
    if st.button(f"🏗️ สร้าง Appendix ({total} สไลด์)", type="primary", use_container_width=True):
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

        errors = []
        for i, fid in enumerate(order):
            f = file_map.get(fid, {})
            progress.progress((i + 1) / total, text=f"({i+1}/{total}) {f.get('name', '')}")
            try:
                img_bytes = download_drive_image(fid, f["api_key"])
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                add_full_image_16_9(slide, img_bytes, prs)
            except Exception:
                errors.append(f.get("name", fid))

        progress.empty()
        if errors:
            st.warning(f"ข้าม {len(errors)} รูป: {', '.join(errors)}")

        output = io.BytesIO()
        prs.save(output)
        st.success(f"✅ สร้างสำเร็จ {total - len(errors)} สไลด์!")
        st.download_button(
            "📥 ดาวน์โหลด Appendix.pptx",
            output.getvalue(),
            "Appendix_16_9.pptx",
            use_container_width=True,
        )
