# ... (ส่วนการประมวลผลไฟล์เหมือนเดิม) ...

if uploaded_file:
    master_obj, slides_data = process_file_optimized(uploaded_file.getvalue(), uploaded_file.name)

    st.header("2. เลือกหน้า (ขยายภาพให้ใหญ่ขึ้น)")
    
    # ปรับเหลือ 2 หรือ 3 คอลัมน์เพื่อให้ภาพใหญ่พอที่จะอ่านรหัสได้
    cols_per_row = 2 
    cols = st.columns(cols_per_row)

    for idx, s in enumerate(slides_data):
        with cols[idx % cols_per_row]:
            # ใส่กรอบเพื่อให้ดูเป็นระเบียบ
            with st.container(border=True):
                if s["preview"]:
                    # แสดงภาพโดยใช้ความกว้างเต็มคอลัมน์
                    st.image(s["preview"], use_container_width=True, caption=s["display"])
                
                # แยก Checkbox ออกมาให้ติ๊กง่ายๆ
                is_checked = idx in st.session_state.selected_slides
                if st.checkbox(f"เลือก {s['display']}", key=f"cb_{idx}", value=is_checked):
                    st.session_state.selected_slides.add(idx)
                else:
                    st.session_state.selected_slides.discard(idx)
                
                # เพิ่มปุ่ม "ดูภาพขยาย" เพื่ออ่านรหัสชัดๆ
                with st.expander("🔍 ดูภาพขยายเพื่ออ่านรหัส"):
                    # ดึงภาพที่ชัดกว่าเดิมมาโชว์ (ถ้าเป็น PDF ให้ดึงใหม่แบบชัดขึ้น)
                    if s["type"] == "pdf":
                        page = master_obj.load_page(s["id"])
                        full_pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                        st.image(io.BytesIO(full_pix.tobytes("png")), use_container_width=True)
                    else:
                        st.image(s["preview"], use_container_width=True)

# ... (ส่วนการสร้างไฟล์เหมือนเดิม) ...
