import streamlit as st
import os
import pandas as pd
import zipfile
import io
import base64
import tempfile 
from PIL import Image
from docx import Document
import fitz 

from core.config import API_KEY, TEMPLATE_DIR
from core.utils.helpers import doc_so_tien_vn, get_tags_from_template
from core.ai_extractor import AIExtractor
from core.document_builder import DocumentBuilder
from ui.components import display_zalo_message

st.set_page_config(page_title="Hệ thống ERP AI", page_icon="🧾", layout="wide")
available_templates = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(('.xlsx', '.docx'))]

st.title("🧾 HỆ THỐNG KẾ TOÁN AI - TẠO BỘ HỒ SƠ TỰ ĐỘNG")
st.markdown("---")

col_top1, col_top2 = st.columns(2)
with col_top1:
    uploaded_file = st.file_uploader("📥 1. Tải lên Hóa đơn / Báo giá", type=["png", "jpg", "jpeg", "xlsx", "docx", "pdf"])
with col_top2:
    if available_templates:
        selected_templates = st.multiselect("📂 2. Chọn Các Biểu Mẫu Muốn Xuất:", available_templates)
    else:
        st.error("⚠️ Chưa có file mẫu nào.")
        selected_templates = []

st.divider()

# Chia tỷ lệ 5:5 cân đối
col_left, col_right = st.columns([5, 5], gap="large")

with col_left:
    st.markdown("### 🔍 TÀI LIỆU GỐC")
    if uploaded_file:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        
        # Tạo file tạm trong thư mục Temp của Hệ điều hành
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_ext}") as tmp_file:
            tmp_file.write(uploaded_file.getbuffer())
            temp_input = tmp_file.name  
            
        with st.container(border=True):
            if file_ext in ['png', 'jpg', 'jpeg']: 
                st.image(Image.open(temp_input), use_container_width=True)
            elif file_ext == 'pdf':
                try:
                    doc = fitz.open(temp_input)
                    st.info(f"📄 Tài liệu PDF có tổng cộng {len(doc)} trang.")
                    for page_num in range(len(doc)):
                        page = doc.load_page(page_num)
                        pix = page.get_pixmap(dpi=150)
                        img_path = f"temp_page_{page_num}.png"
                        pix.save(img_path)
                        st.image(Image.open(img_path), caption=f"Trang {page_num + 1}", use_container_width=True)
                        if os.path.exists(img_path):
                            os.remove(img_path)
                except Exception as e:
                    st.error(f"Không thể hiển thị bản xem trước PDF: {str(e)}")
                    with open(temp_input, "rb") as f:
                        st.download_button("📥 Tải xuống file PDF gốc để xem", f, file_name=uploaded_file.name)
            elif file_ext == 'xlsx':
                for sheet, df in pd.read_excel(temp_input, sheet_name=None).items(): 
                    st.dataframe(df, height=800, use_container_width=True)
            elif file_ext == 'docx':
                text = "\n".join([p.text for p in Document(temp_input).paragraphs])
                st.text_area("Nội dung Word:", text, height=850)

with col_right:
    st.markdown(f"### ✍️ KIỂM DUYỆT THÔNG TIN")
    if uploaded_file and selected_templates:
        
        all_dynamic_tags = set()
        for tpl in selected_templates:
            tpl_path = os.path.join(TEMPLATE_DIR, tpl)
            all_dynamic_tags.update(get_tags_from_template(tpl_path))
            
        if st.button("🚀 KÍCH HOẠT AI ĐỌC MẪU & ĐỐI CHIẾU", type="primary", use_container_width=True):
            with st.spinner("AI đang tính toán và lấy dữ liệu..."):
                data = AIExtractor(API_KEY).extract_invoice_data(temp_input, expected_tags=all_dynamic_tags) 
                if "error" not in data: st.session_state['data'] = data
                else: st.error(data['error'])

        if 'data' in st.session_state:
            data = st.session_state['data']
            
            canh_bao = data.get("danh_sach_canh_bao", [])
            if canh_bao and len(canh_bao) > 0 and canh_bao[0].strip() != "":
                st.error("🚨 **CẢNH BÁO TỪ AI!**")
                for loi in canh_bao: st.warning(f"⚠️ {loi}")
            else:
                st.success("✅ AI báo cáo: Số liệu chuẩn khớp.")

            ai_flat_data = {}
            ai_flat_data.update(data.get('thong_tin_nha_cung_cap', {}))
            ai_flat_data.update(data.get('thong_tin_chung', {}))
            ai_flat_data.update(data.get('thong_tin_khach_hang', {}))
            ai_flat_data.update(data.get('thong_tin_dong', {})) 
            
            final_dynamic_data = {} 
            
            if all_dynamic_tags:
                with st.container(border=True):
                    st.markdown("**Thông Tin Báo Cáo (Tự động quét từ Hóa đơn):**")
                    cols = st.columns(3) 
                    
                    for idx, tag in enumerate(sorted(all_dynamic_tags)):
                        col = cols[idx % 3]
                        default_val = str(ai_flat_data.get(tag, "")) 
                        label = tag.replace("_", " ").title() 
                        with col:
                            final_dynamic_data[tag] = st.text_input(label, value=default_val, key=f"dyn_{tag}")
            
            st.markdown("**Bảng Kê Hàng Hóa / Dịch Vụ:**")
            ds = data.get("danh_sach_hang_hoa", [])
            if not ds: ds = [{"ten_hang_hoa": ""}]
            
            fmt_data = []
            for item in ds:
                gia = float(item.get("don_gia") or 0)
                sl = float(item.get("so_luong") or 0)
                fmt_data.append({
                    "Tên hàng hóa": item.get("ten_hang_hoa", ""), "Mục đích SD": item.get("muc_dich_su_dung", ""), 
                    "Tồn kho": "", "Nhà CC": ai_flat_data.get('ten_cong_ty', ''),
                    "ĐVT": item.get("don_vi_tinh", ""), "Số lượng": sl, "Đơn giá": gia, 
                    "Thành tiền": item.get("thanh_tien", sl * gia), "Ghi chú": item.get("ghi_chu", "")
                })
            
            edited_df = st.data_editor(pd.DataFrame(fmt_data), num_rows="dynamic", use_container_width=True, hide_index=True)
            danh_sach_da_chinh_sua = edited_df.to_dict('records')

            ds_chuan = []
            for item in danh_sach_da_chinh_sua:
                ten_hh = str(item.get("Tên hàng hóa") or "").strip()
                if not ten_hh or ten_hh == "None": continue 
                ds_chuan.append({
                    "ten_hang_hoa": ten_hh, "muc_dich": str(item.get("Mục đích SD") or ""),
                    "ton_kho": str(item.get("Tồn kho") or ""), "nha_cung_cap": str(item.get("Nhà CC") or ""),
                    "don_vi_tinh": str(item.get("ĐVT") or ""), "so_luong": float(item.get("Số lượng") or 0),
                    "don_gia": float(item.get("Đơn giá") or 0), "thanh_tien": float(item.get("Thành tiền") or 0), 
                    "ghi_chu": str(item.get("Ghi chú") or "")
                })

            tong_tien_hang = sum(item.get("thanh_tien", 0) for item in ds_chuan)
            
            st.markdown("---")
            col_t1, col_t2, col_t3 = st.columns(3)
            with col_t1: st.metric("Cộng tiền hàng:", f"{tong_tien_hang:,.0f} đ")
            with col_t2: 
                thue_suat = st.number_input("Thuế suất GTGT (%)", value=8)
                tien_thue = tong_tien_hang * (thue_suat / 100)
                st.metric("Tiền thuế GTGT:", f"{tien_thue:,.0f} đ")
            with col_t3:
                tong_thanh_toan = tong_tien_hang + tien_thue
                st.metric("Tổng cộng thanh toán:", f"{tong_thanh_toan:,.0f} đ")
                chu_so_tien = doc_so_tien_vn(tong_thanh_toan)
                st.markdown(f"*(Bằng chữ: {chu_so_tien})*")

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("✅ XUẤT BỘ HỒ SƠ CHUẨN ERP", type="primary", use_container_width=True):
                
                for item in ds_chuan:
                    tong_cong_mon = item["thanh_tien"] * (1 + (thue_suat / 100))
                    item["tong_cong"] = f"{tong_cong_mon:,.0f}"
                
                final_data = {
                    **final_dynamic_data, 
                    "tong_tien_hang": tong_tien_hang, "thue_gtgt": tien_thue, 
                    "tong_thanh_toan": tong_thanh_toan, "tong_cong": tong_thanh_toan,
                    "so_tien_bang_chu": chu_so_tien, 
                    "danh_sach_hang_hoa": ds_chuan
                }
                
                try:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for tpl in selected_templates:
                            tpl_path = os.path.join(TEMPLATE_DIR, tpl)
                            out_file = f"BaoCao_{tpl}"
                            
                            if tpl.endswith('.docx'): DocumentBuilder.fill_word_template(tpl_path, out_file, final_data)
                            else: DocumentBuilder.fill_excel_template(tpl_path, out_file, final_data)
                            
                            zip_file.write(out_file, arcname=out_file)
                            os.remove(out_file) 
                    
                    st.success("Tuyệt vời! Toàn bộ file đã được xử lý xong.")
                    st.download_button(
                        label="💾 BẤM VÀO ĐÂY ĐỂ TẢI BỘ HỒ SƠ (FILE ZIP)", 
                        data=zip_buffer.getvalue(), file_name="Bo_Ho_So_Ketoan.zip", 
                        mime="application/zip", type="secondary", use_container_width=True
                    )
                    
                    ten_doi_tac = final_dynamic_data.get('ten_cong_ty_ben_b', ai_flat_data.get('ten_cong_ty', ''))
                    display_zalo_message(ten_doi_tac, tong_thanh_toan)

                except Exception as e:
                    st.error(f"Lỗi: {str(e)}")
