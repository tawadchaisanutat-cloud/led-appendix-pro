import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import requests
from PIL import Image
from streamlit_sortables import sort_items 

# --- 1. CONFIG ---
st.set_page_config(page_title="LED Appendix Pro", layout="wide")

secrets = st.secrets if hasattr(st, 'secrets') else {}
API_KEY = secrets.get("GDRIVE_API_KEY", "")
FOLDER_ID = secrets.get("GDRIVE_FOLDER_ID", "")

# --- 2. SESSION STATE ---
if 'selected_images' not in st.session_state: st.session_state.selected_images = set()
if 'image_order' not in st.session_state: st.session_state.image_order = []
if 'drive_files' not in st.session_state: st.session_state.drive_files = []

# --- 3. FUNCTIONS ---
@st.cache_data(show_spinner="🚀 โหลดข้อมูล...", ttl=3600)
def fetch_all_files(api_key, folder_id):
    if not api_key or not folder_id: return []
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
    except: return []

@st.cache_data(show_spinner=False)
def download_image(file_id, api_key):
    try:
        r = requests.get(f"https://www.googleapis.com/drive/v3/files/{file_id}",
                         params={"alt": "media", "key": api_key}, timeout=30)
        return r.content
    except: return None

def add_image_to_slide(slide, img_bytes, prs, is_cover=False):
    if not img_bytes: return
    try:
        if is_cover:
            # ขยายปกให้เต็มหน้าจอ 16:9
            slide.shapes.add_picture(io.BytesIO(img_bytes), 0, 0, width=prs.slide_width, height=prs.slide_height)
        else:
            img = Image.open(io.BytesIO(img_bytes))
            w, h = img.size
            ratio = min(prs.slide_width/w, prs.slide_height/h)
            nw, nh = int(w*ratio), int(h*ratio)
            slide.shapes.add_picture(io.BytesIO(img_bytes), (prs.slide_width-nw)//2, (prs.slide_height-nh)//2, width=nw, height=nh)
    except: pass

# --- 4. ⚡️ AUTO-LOAD (บังคับรันทันที) ---
if API_KEY and FOLDER_ID and not st.session_state.drive_files:
    st.session_state.drive_files = fetch_all_files(API_KEY, FOLDER_ID)
    if st.session_state.drive_files: st.rerun()

# --- 5. UI CONTENT ---
st.title("🚀 LED Appendix Pro")
all_files = st.session_state.drive_files

if not all_files:
    st.error("❌ ไม่พบข้อมูล! โปรดเช็ก Secrets")
    if st.sidebar.button("🔄 ลองโหลดอีกครั้ง"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

# --- ค้นหาปกอิงตามชื่อไฟล์ในรูปของคุณ (COVER A / COVER B) ---
def find_cover(key):
    # ค้นหาแบบไม่สนเว้นวรรคหรือตัวเล็กใหญ่
    for f in all_files:
        clean_name = f['name'].upper().replace(" ", "")
        if key in clean_name: return f
    return None

cover_a = find_cover("COVERA")
cover_b = find_cover("COVERB")

with st.expander("📝 ตั้งค่าข้อความบนหน้าปก", expanded=False):
    c1, c2 = st.columns(2)
    p_title = c1.text_input("หัวข้อโปรเจกต์", "PROPOSAL FOR DIGITAL LED")
    p_sub = c2.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")

tab1, tab2 = st.tabs(["📷 เลือกรูปภาพ", f"🔀 จัดเรียง ({len(st.session_state.selected_images)})"])

# --- TAB 1: เลือกรูป ---
with tab1:
    search = st.text_input("🔍 ค้นหารหัสภาพ", placeholder="พิมพ์ชื่อภาพ...").upper()
    display = [f for f in all_files if (search and search in f["name"].upper()) or (f["id"] in st.session_state.selected_images)]
    
    if not display: st.info("💡 พิมพ์รหัสภาพเพื่อค้นหา")
    else:
        cols = st.columns(4)
        for i, f in enumerate(display):
            with cols[i % 4]:
                with st.container(border=True):
                    st.image(f.get("thumbnailLink", "").replace("=s220", "=s400"), use_container_width=True)
                    fid, fname = f["id"], f["name"]
                    is_sel = fid in st.session_state.selected_images
                    if st.checkbox(fname[:15], key=f"sel_{fid}", value=is_sel):
                        if fid not in st.session_state.selected_images:
                            st.session_state.selected_images.add(fid)
                            st.session_state.image_order.append(fid)
                    else:
                        if fid in st.session_state.selected_images:
                            st.session_state.selected_images.discard(fid)
                            st.session_state.image_order = [x for x in st.session_state.image_order if x != fid]

# --- TAB 2: จัดเรียง (เห็นภาพชัดเจน) ---
with tab2:
    valid_order = [fid for fid in st.session_state.image_order if fid in st.session_state.selected_images]
    if not valid_order: st.info("เลือกรูปภาพก่อนครับ")
    else:
        st.subheader("🖱️ เรียงลำดับ (ลากแถบเพื่อสลับที่)")
        file_map = {f['id']: f for f in all_files}
        sort_input = [f"{file_map[fid]['name']} | {fid}" for fid in valid_order]
        
        # แก้ ValueError ด้วยการส่ง List[str] ล้วนๆ
        sorted_output = sort_items(sort_input, direction="vertical", key="sort_v_final")
        st.session_state.image_order = [s.split(" | ")[-1] for s in sorted_output]
        
        # แสดง Preview รูปที่เรียงแล้วให้ดูทันที
        st.markdown("**🔍 ลำดับที่จะปรากฏในไฟล์ (เรียงตามสไลด์):**")
        p_cols = st.columns(6)
        for i, fid in enumerate(st.session_state.image_order):
            with p_cols[i % 6]:
                st.image(file_map[fid].get("thumbnailLink", ""), caption=f"ลำดับ {i+1}")

        if st.button("🏗️ สร้างไฟล์ PowerPoint (.pptx)", type="primary", use_container_width=True):
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            
            # 1. ใส่ปกหน้า (ถ้าเจอ COVER A)
            slide_a = prs.slides.add_slide(prs.slide_layouts[6])
            if cover_a:
                add_image_to_slide(slide_a, download_image(cover_a['id'], API_KEY), prs, is_cover=True)
            
            # วาดข้อความทับปก
            tb = slide_a.shapes.add_textbox(0, prs.slide_height/2.5, prs.slide_width, Inches(3))
            p = tb.text_frame.paragraphs[0]
            p.text, p.alignment, p.font.size, p.font.bold = p_title, PP_ALIGN.CENTER, Pt(50), True
            if cover_a: p.font.color.rgb = RGBColor(255, 255, 255)
            
            p2 = tb.text_frame.add_paragraph()
            p2.text, p2.alignment, p2.font.size = p_sub, PP_ALIGN.CENTER, Pt(28)
            if cover_a: p2.font.color.rgb = RGBColor(255, 255, 255)

            # 2. ใส่รูปเนื้อหา (ข้ามรูปที่เป็นปก)
            for fid in st.session_state.image_order:
                if (cover_a and fid == cover_a['id']) or (cover_b and fid == cover_b['id']): continue
                add_image_to_slide(prs.slides.add_slide(prs.slide_layouts[6]), download_image(fid, API_KEY), prs)

            # 3. ใส่ปกหลัง (ถ้าเจอ COVER B)
            if cover_b:
                slide_b = prs.slides.add_slide(prs.slide_layouts[6])
                add_image_to_slide(slide_b, download_image(cover_b['id'], API_KEY), prs, is_cover=True)
            
            out = io.BytesIO()
            prs.save(out)
            st.success("🎉 Export สำเร็จ!")
            st.download_button("📥 ดาวน์โหลด Appendix", out.getvalue(), "Appendix.pptx", use_container_width=True)

# --- 6. SIDEBAR STATUS ---
with st.sidebar:
    st.subheader("📊 สถานะปกอัตโนมัติ")
    st.write(f"🖼️ ปกหน้า (COVER A): {'✅ พบแล้ว' if cover_a else '❌ ไม่พบ'}")
    st.write(f"🖼️ ปกหลัง (COVER B): {'✅ พบแล้ว' if cover_b else '❌ ไม่พบ'}")
    if st.button("🗑️ ล้าง Cache และโหลดใหม่"):
        st.cache_data.clear()
        st.session_state.drive_files = []
        st.rerun()
