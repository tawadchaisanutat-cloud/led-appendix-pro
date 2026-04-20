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
st.set_page_config(page_title="LED Appendix Pro", layout="wide")

# --- 2. GET SECRETS (บังคับดึงค่าทันที) ---
secrets = st.secrets if hasattr(st, 'secrets') else {}
API_KEY = secrets.get("GDRIVE_API_KEY", "")
FOLDER_ID = secrets.get("GDRIVE_FOLDER_ID", "")

# --- 3. CORE FUNCTIONS (มี Cache แต่บังคับทำงาน) ---
@st.cache_data(show_spinner="🚀 กำลังดึงฐานข้อมูลรูปภาพ...", ttl=3600)
def get_data_immediately(api_key, folder_id):
    """ฟังก์ชันนี้ต้องคืนค่าข้อมูลเสมอถ้ามี API Key"""
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
def download_image_final(file_id, api_key):
    r = requests.get(f"https://www.googleapis.com/drive/v3/files/{file_id}",
                     params={"alt": "media", "key": api_key}, timeout=30)
    return r.content

# --- 4. ⚡️ THE "NO-SKIP" LOAD (บังคับรันก่อนขึ้น UI) ---
# บรรทัดนี้จะรันทุกครั้งที่เปิดแอป และจะดึงข้อมูลมารอไว้ในตัวแปร all_files ทันที
all_files = get_data_immediately(API_KEY, FOLDER_ID)

# เริ่มต้น State (ต้องอยู่หลังการโหลดข้อมูล)
if 'selected_images' not in st.session_state: st.session_state.selected_images = set()
if 'image_order' not in st.session_state: st.session_state.image_order = []

# --- 5. UI: MAIN SCREEN ---
st.title("🚀 LED Appendix Pro")

# ถ้าข้อมูลยังไม่มา (เช่น API ผิด) ให้หยุดแค่ตรงนี้
if not all_files:
    st.error("❌ ไม่พบข้อมูลใน Google Drive! โปรดเช็ก Secrets (GDRIVE_API_KEY / GDRIVE_FOLDER_ID)")
    if st.button("ลองโหลดใหม่อีกครั้ง"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

# --- 6. APP CONTENT (จะทำงานก็ต่อเมื่อมีข้อมูลแล้วเท่านั้น) ---

# ส่วนตั้งค่าหน้าปก
with st.expander("📝 ตั้งค่าหน้าปกและปกหลัง", expanded=False):
    c1, c2 = st.columns(2)
    c_title = c1.text_input("หัวข้อโปรเจกต์", "PROPOSAL FOR DIGITAL LED")
    c_sub = c2.text_input("ชื่อลูกค้า", "Presented to: Valued Customer")

# ค้นหา CoverA/B อัตโนมัติ
cover_a = next((f for f in all_files if "COVERA" in f['name'].upper()), None)
cover_b = next((f for f in all_files if "COVERB" in f['name'].upper()), None)

tab1, tab2 = st.tabs(["📷 เลือกรูปภาพ", f"🔀 จัดเรียง ({len(st.session_state.selected_images)})"])

with tab1:
    # ช่อง Search จะทำงานได้ทันทีเพราะข้อมูลอยู่ใน all_files แล้ว
    search = st.text_input("🔍 ค้นหารหัสภาพ", placeholder="พิมพ์รหัสภาพที่นี่ รูปจะขึ้นทันที...").upper()
    
    # Filter: โชว์เฉพาะที่ Search เจอ + ที่ติ๊กเลือกไว้
    display_files = [f for f in all_files if (search and search in f["name"].upper()) or (f["id"] in st.session_state.selected_images)]
    
    if not display_files:
        st.info("💡 พิมพ์รหัสภาพเพื่อเริ่มค้นหา (รูปที่ถูกเลือกไว้จะแสดงค้างไว้ที่นี่เสมอ)")
    else:
        cols = st.columns(4)
        for idx, f in enumerate(display_files):
            with cols[idx % 4]:
                with st.container(border=True):
                    # Thumbnail จาก Google (แก้ตัวแปรให้ตรงเป๊ะ)
                    t_url = f.get("thumbnailLink", "").replace("=s220", "=s400")
                    if t_url: st.image(t_url, use_container_width=True)
                    
                    fid, fname = f["id"], f["name"]
                    is_sel = fid in st.session_state.selected_images
                    
                    if st.checkbox(fname[:15], key=f"ch_{fid}", value=is_sel):
                        if fid not in st.session_state.selected_images:
                            st.session_state.selected_images.add(fid)
                            st.session_state.image_order.append(fid)
                    else:
                        if fid in st.session_state.selected_images:
                            st.session_state.selected_images.discard(fid)
                            st.session_state.image_order = [x for x in st.session_state.image_order if x != fid]

with tab2:
    # กรองเฉพาะ ID ที่ยังคงเลือกอยู่
    current_order = [fid for fid in st.session_state.image_order if fid in st.session_state.selected_images]
    
    if not current_order:
        st.info("กรุณาเลือกรูปภาพจากแท็บแรกก่อนครับ")
    else:
        st.subheader("🖱️ คลิกลากเพื่อเรียงลำดับ (บนสุด = สไลด์แรก)")
        id_to_name = {f['id']: f['name'] for f in all_files}
        
        # --- FIX VALUEERROR: บังคับข้อมูลให้เป็น List ของ String ล้วนๆ ---
        sort_input = [f"{id_to_name.get(fid, 'Unknown')} | {fid}" for fid in current_order]
        
        # แสดงตัวลากวาง
        sorted_output = sort_items(sort_input, direction="vertical", key="stable_sort_v7")
        
        # แกะ ID กลับออกมา
        new_id_order = [item.split(" | ")[-1] for item in sorted_output]
        st.session_state.image_order = new_id_order

        st.divider()

        if st.button("🏗️ สร้างไฟล์ PowerPoint (.pptx)", type="primary", use_container_width=True):
            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
            
            # --- ฟังก์ชันช่วยวาดรูป (ประกาศไว้ข้างในเพื่อความชัวร์) ---
            def draw_img(slide, img_bytes):
                try:
                    img_io = io.BytesIO(img_bytes)
                    img_open = Image.open(img_io)
                    w, h = img_open.size
                    ratio = min(prs.slide_width/w, prs.slide_height/h)
                    nw, nh = int(w*ratio), int(h*ratio)
                    left, top = (prs.slide_width-nw)//2, (prs.slide_height-nh)//2
                    slide.shapes.add_picture(img_io, left, top, width=nw, height=nh)
                except: pass

            # 1. ปกหน้า
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            if cover_a:
                draw_img(slide, download_image_final(cover_a['id'], API_KEY))
            tb = slide.shapes.add_textbox(0, prs.slide_height/2.5, prs.slide_width, Inches(2))
            p = tb.text_frame.paragraphs[0]
            p.text, p.alignment, p.font.size, p.font.bold = c_title, PP_ALIGN.CENTER, Pt(48), True
            if cover_a: p.font.color.rgb = RGBColor(255, 255, 255)
            
            # 2. เนื้อหา
            for fid in new_id_order:
                if (cover_a and fid == cover_a['id']) or (cover_b and fid == cover_b['id']): continue
                draw_img(prs.slides.add_slide(prs.slide_layouts[6]), download_image_final(fid, API_KEY))
            
            # 3. ปกหลัง
            if cover_b:
                draw_img(prs.slides.add_slide(prs.slide_layouts[6]), download_image_final(cover_b['id'], API_KEY))
            
            out = io.BytesIO()
            prs.save(out)
            st.success("🎉 สำเร็จ!")
            st.download_button("📥 ดาวน์โหลดไฟล์", out.getvalue(), "Appendix.pptx", use_container_width=True)

# --- 7. SIDEBAR ---
with st.sidebar:
    st.write(f"✅ พบข้อมูลใน Drive: {len(all_files)} รูป")
    if st.button("🔄 ล้าง Cache และโหลดใหม่"):
        st.cache_data.clear()
        st.rerun()
