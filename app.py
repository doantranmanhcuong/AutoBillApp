import streamlit as st
import os
import pandas as pd
import zipfile
import io
import tempfile
import time
from PIL import Image
from docx import Document
try:
    import pymupdf as fitz
except ImportError:
    import fitz

from core.config import API_KEY, API_KEYS, TEMPLATE_DIR
from core.utils.helpers import doc_so_tien_vn, exhaustive_extract_tags
from core.ai_extractor import AIExtractor
from core.invoice_merger import InvoiceMerger
from core.document_builder import DocumentBuilder
from ui.components import apply_office_theme, display_financial_summary, display_zalo_message

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Hệ thống Hồ sơ Kế toán Tự động",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Nạp giao diện văn phòng chuẩn mực (đơn giản, trong sáng)
apply_office_theme()

def render_file_preview(file_name: str, file_path: str, unique_idx: int = 0):
    """Hiển thị nội dung xem trước tài liệu gốc (PDF, Ảnh, Excel, Word)"""
    file_ext = file_name.split('.')[-1].lower()
    with st.container(border=True):
        if file_ext in ['png', 'jpg', 'jpeg']: 
            st.image(file_path, use_container_width=True)
        elif file_ext == 'pdf':
            try:
                doc = fitz.open(file_path)
                st.caption(f"📄 Tài liệu PDF gồm {len(doc)} trang:")
                for page_num in range(len(doc)):
                    page = doc.load_page(page_num)
                    pix = page.get_pixmap(dpi=150)
                    img_bytes = pix.tobytes("png")
                    st.image(img_bytes, caption=f"Trang {page_num + 1}", use_container_width=True)
            except Exception as e:
                st.warning(f"Không thể mở trực tiếp bản xem trước PDF: {str(e)}")
                with open(file_path, "rb") as f:
                    st.download_button("📥 Tải về file PDF gốc để xem", f, file_name=file_name, key=f"dl_{unique_idx}_{file_name}")
        elif file_ext == 'xlsx':
            for sheet, df in pd.read_excel(file_path, sheet_name=None).items(): 
                st.caption(f"Trang tính: {sheet}")
                clean_df = df.fillna("").astype(str)
                st.dataframe(clean_df, height=500, use_container_width=True)
        elif file_ext == 'docx':
            text = "\n".join([p.text for p in Document(file_path).paragraphs])
            st.text_area("Nội dung văn bản Word:", text, height=500, key=f"txt_{unique_idx}_{file_name}")

# Tiêu đề ứng dụng
st.markdown("### 📑 HỆ THỐNG LẬP HỒ SƠ & CHỨNG TỪ KẾ TOÁN")
st.caption("Giải pháp tự động hóa lập chứng từ, đề nghị thanh toán và hợp đồng từ hóa đơn / báo giá")
st.markdown("---")

# Kiểm tra thư mục biểu mẫu
available_templates = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(('.xlsx', '.docx'))]

# BƯỚC 1: KHỞI TẠO HỒ SƠ
col_top1, col_top2 = st.columns([1, 1], gap="medium")
with col_top1:
    uploaded_files = st.file_uploader(
        "📁 1. Tải lên Hóa đơn / Báo giá (Cho phép chọn nhiều hóa đơn cùng NCC):", 
        type=["png", "jpg", "jpeg", "xlsx", "docx", "pdf"],
        accept_multiple_files=True
    )
with col_top2:
    if available_templates:
        selected_templates = st.multiselect(
            f"📂 2. Chọn các biểu mẫu cần xuất ({len(available_templates)} mẫu có sẵn):", 
            available_templates
        )
    else:
        st.warning("⚠️ Thư mục 'templates' chưa có biểu mẫu (.docx hoặc .xlsx nào).")
        selected_templates = []

st.markdown("---")

# Quản lý danh sách file tạm an toàn theo list (tránh trùng tên và bảo toàn thứ tự)
current_signatures = [f"{f.name}_{f.size}" for f in uploaded_files] if uploaded_files else []
if st.session_state.get('uploaded_signatures') != current_signatures:
    old_temp_list = st.session_state.get('temp_files_list', [])
    for item in old_temp_list:
        p = item.get("path")
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass
    
    new_temp_list = []
    for f in (uploaded_files or []):
        f_ext = f.name.split('.')[-1].lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=f".{f_ext}") as tmp_f:
            tmp_f.write(f.getbuffer())
            new_temp_list.append({"name": f.name, "path": tmp_f.name})
            
    st.session_state['temp_files_list'] = new_temp_list
    st.session_state['uploaded_signatures'] = current_signatures
    if 'data' in st.session_state:
        del st.session_state['data']

temp_files_list = st.session_state.get('temp_files_list', [])

# BƯỚC 2 & 3: GIAO DIỆN 2 CỘT (CHỨNG TỪ GỐC & KIỂM DUYỆT)
col_left, col_right = st.columns([4.8, 5.2], gap="large")

# CỘT TRÁI: HIỂN THỊ CHỨNG TỪ GỐC
with col_left:
    st.markdown("##### 🔍 TÀI LIỆU GỐC")
    if not uploaded_files or not temp_files_list:
        st.info("Vui lòng tải lên tài liệu ở Bước 1 để bắt đầu xem trước.")
    elif len(temp_files_list) == 1:
        item = temp_files_list[0]
        render_file_preview(item["name"], item["path"], unique_idx=0)
    else:
        st.caption(f"📑 Đã nạp {len(temp_files_list)} hóa đơn (Chuyển tab để xem từng file):")
        file_tabs = st.tabs([f"📄 {item['name']}" for item in temp_files_list])
        for idx, item in enumerate(temp_files_list):
            with file_tabs[idx]:
                render_file_preview(item["name"], item["path"], unique_idx=idx)

# CỘT PHẢI: KIỂM DUYỆT THÔNG TIN & ĐỐI CHIẾU
with col_right:
    st.markdown("##### ✍️ KIỂM DUYỆT & ĐỐI CHIẾU THÔNG TIN")
    
    if not uploaded_files:
        st.info("Chưa có chứng từ nào được nạp.")
    elif not selected_templates:
        st.info("Vui lòng chọn ít nhất 1 biểu mẫu ở ô trên để hệ thống nạp các trường cần điền.")
    else:
        # Quét các tag cần trích xuất từ các biểu mẫu đã chọn
        template_tags_map = {}
        all_dynamic_tags = set()
        
        TABLE_TAGS = [
            "stt", "ten_hang_hoa", "don_vi_tinh", "so_luong", "don_gia", "thanh_tien", "tong_cong",
            "tong_tien_hang", "thue_gtgt", "tong_thanh_toan", "so_tien_bang_chu", 
            "ghi_chu", "muc_dich_su_dung", "ton_kho", "nha_cung_cap", "muc_dich", "tr"
        ]
        
        for tpl in selected_templates:
            tpl_path = os.path.join(TEMPLATE_DIR, tpl)
            raw_tags = exhaustive_extract_tags(tpl_path)
            clean_tags = set()
            for tag in raw_tags:
                t = str(tag).lower().strip()
                if not ("%" in t or t.startswith("item.") or "endfor" in t or "for " in t or t in TABLE_TAGS):
                    clean_tags.add(t)
                    all_dynamic_tags.add(t)
            template_tags_map[tpl] = clean_tags
            
        # Nút kích hoạt trích xuất
        if not API_KEYS:
            st.error("Chưa tìm thấy GEMINI_API_KEYS trong file .env hoặc Secrets. Vui lòng kiểm tra lại cấu hình.")
        else:
            btn_label = "🔍 Trích xuất thông tin chứng từ" if len(temp_files_list) == 1 else f"🔍 Trích xuất & Gộp {len(temp_files_list)} hóa đơn"
            if st.button(btn_label, type="primary", use_container_width=True):
                invoice_results = []
                progress_bar = st.progress(0.0)
                status_box = st.empty()
                total_files = len(temp_files_list)
                extractor = AIExtractor(API_KEYS)

                for idx, item in enumerate(temp_files_list):
                    fname = item["name"]
                    fpath = item["path"]

                    # Giãn cách 2.0 giây giữa các file để tránh vượt giới hạn tốc độ RPM của AI
                    if idx > 0:
                        status_box.caption(f"⏳ Tạm dừng 2s để làm nguội máy chủ AI trước khi bóc tách file {idx + 1}/{total_files}...")
                        time.sleep(2.0)

                    status_box.info(f"⏳ Đang bóc tách hóa đơn {idx + 1}/{total_files}: **{fname}**...")
                    progress_bar.progress(idx / total_files)
                    
                    extracted = extractor.extract_invoice_data(fpath, expected_tags=all_dynamic_tags)
                    invoice_results.append({
                        "filename": fname,
                        "data": extracted
                    })

                progress_bar.progress(1.0)
                status_box.empty()
                progress_bar.empty()

                if len(invoice_results) == 1:
                    merged_res = invoice_results[0]["data"]
                else:
                    merged_res = InvoiceMerger.merge_invoices(invoice_results)

                if "error" not in merged_res:
                    st.session_state['data'] = merged_res
                    meta = merged_res.get("_multi_invoice_meta", {})
                    failed_cnt = meta.get("failed_count", 0)
                    if failed_cnt > 0:
                        st.warning(f"⚠️ Đã gộp {meta.get('total_invoices')}/{meta.get('total_uploaded')} hóa đơn. Có {failed_cnt} hóa đơn bị lỗi: {', '.join(meta.get('failed_files', []))}")
                    else:
                        success_msg = "Trích xuất thông tin chứng từ thành công!" if len(temp_files_list) == 1 else f"Đã trích xuất & gộp trọn vẹn cả {len(temp_files_list)} hóa đơn thành công!"
                        st.toast(success_msg, icon="✅")
                else:
                    st.error(merged_res['error'])

        # Khi đã có dữ liệu trích xuất
        if 'data' in st.session_state:
            data = st.session_state['data']
            
            # Thông báo gộp nhiều hóa đơn nếu có
            meta = data.get("_multi_invoice_meta")
            if meta and meta.get("total_uploaded", 1) > 1:
                ncc_name = data.get("thong_tin_nha_cung_cap", {}).get("ten_cong_ty") or "Nhà cung cấp"
                so_hd_gop = data.get("thong_tin_chung", {}).get("so_chung_tu") or "Đã tổng hợp"
                failed_cnt = meta.get("failed_count", 0)
                if failed_cnt > 0:
                    st.warning(f"📑 **ĐÃ GỘP {meta['total_invoices']}/{meta['total_uploaded']} HÓA ĐƠN (CÓ {failed_cnt} HÓA ĐƠN BỊ LỖI):**\n\n- **Nhà cung cấp:** {ncc_name}\n- **Số hóa đơn:** {so_hd_gop}\n- **Tổng mặt hàng:** {len(data.get('danh_sach_hang_hoa', []))} mục\n- **File bị lỗi:** {', '.join(meta.get('failed_files', []))}")
                else:
                    st.success(f"📑 **ĐÃ GỘP THÀNH CÔNG TRỌN VẸN CẢ {meta['total_invoices']} HÓA ĐƠN!**\n\n- **Nhà cung cấp:** {ncc_name}\n- **Số hóa đơn:** {so_hd_gop}\n- **Tổng mặt hàng:** {len(data.get('danh_sach_hang_hoa', []))} mục")

            # Cảnh báo sai lệch số liệu hoặc khác NCC (nếu có)
            canh_bao = data.get("danh_sach_canh_bao", [])
            if canh_bao and len(canh_bao) > 0 and str(canh_bao[0]).strip() != "":
                st.warning("⚠️ **Lưu ý số liệu & Nhà cung cấp:**")
                for loi in canh_bao: 
                    st.caption(f"- {loi}")

            # Trải phẳng dữ liệu (Flatten)
            ai_flat_data = {}
            if isinstance(data, dict):
                for k, v in data.items():
                    if isinstance(v, dict):
                        for sub_k, sub_v in v.items():
                            ai_flat_data[sub_k] = sub_v
                    elif not isinstance(v, list):
                        ai_flat_data[k] = v
            
            final_dynamic_data = {} 
            
            # 1. KIỂM DUYỆT THÔNG TIN CHUNG TỪNG BIỂU MẪU
            if template_tags_map:
                st.markdown("**1. Thông tin chung theo biểu mẫu:**")
                
                TAG_MAPPING = {
                    "ho_ten_nguoi_de_nghi": "Họ tên người đề nghị",
                    "bo_phan": "Bộ phận",
                    "bo_phan_cong_tac": "Bộ phận công tác",
                    "ly_do_de_nghi": "Lý do đề nghị",
                    "ly_do_thanh_toan": "Lý do thanh toán",
                    "hinh_thuc_thanh_toan": "Hình thức thanh toán/tạm ứng",
                    "thoi_han_hoan_ung": "Thời hạn hoàn ứng",
                    "ngay_lap_phieu": "Ngày lập phiếu",
                    "ngay_ky": "Ngày ký",
                    "thang_ky": "Tháng ký",
                    "nam_ky": "Năm ký",
                    "ngay_thang_nam": "Ngày tháng năm",
                    "so_phieu": "Số phiếu",
                    "so_hop_dong": "Số hợp đồng",
                    "ten_cong_ty": "Tên công ty",
                    "ten_cong_ty_ben_b": "Tên công ty Bên B",
                    "nguoi_dai_dien_ben_b": "Người đại diện Bên B",
                    "chuc_vu_ben_b": "Chức vụ Bên B",
                    "dia_chi": "Địa chỉ",
                    "dia_chi_ben_b": "Địa chỉ Bên B",
                    "ma_so_thue": "Mã số thuế",
                    "mst_ben_b": "MST Bên B",
                    "dien_thoai": "Điện thoại",
                    "dien_thoai_ben_b": "Điện thoại Bên B",
                    "so_tai_khoan": "Số tài khoản",
                    "so_tai_khoan_ben_b": "Số tài khoản Bên B",
                    "ten_ngan_hang": "Tên ngân hàng",
                    "thoi_gian_thuc_hien": "Thời gian thực hiện",
                    "dia_diem_thuc_hien": "Địa điểm thực hiện",
                    "thoi_gian_bao_hanh": "Thời gian bảo hành",
                    "ty_le_tam_ung": "Tỷ lệ tạm ứng (%)",
                    "so_don_dat_hang": "Số đơn đặt hàng",
                    "so_de_xuat": "Số đề xuất",
                    "nguoi_phu_trach": "Người phụ trách",
                    "email_phu_trach": "Email phụ trách",
                    "nguoi_nhan_hang": "Người nhận hàng",
                    "sdt_nguoi_nhan": "SĐT người nhận",
                    "nguoi_de_xuat": "Người đề xuất"
                }
                
                displayed_tags = set()
                
                for tpl in selected_templates:
                    tags_in_tpl = template_tags_map.get(tpl, set())
                    
                    with st.expander(f"📄 Biểu mẫu: {tpl}", expanded=True):
                        if not tags_in_tpl:
                            st.caption("Biểu mẫu này chỉ sử dụng bảng kê hàng hóa, không có trường thông tin chung riêng.")
                        else:
                            cols = st.columns(2) 
                            for idx, tag in enumerate(sorted(tags_in_tpl)):
                                col = cols[idx % 2]
                                label = TAG_MAPPING.get(tag, tag.replace("_", " ").title())
                                
                                with col:
                                    if tag not in displayed_tags:
                                        default_val = str(ai_flat_data.get(tag, "")) 
                                        final_dynamic_data[tag] = st.text_input(
                                            label, 
                                            value=default_val, 
                                            key=f"dyn_{tag}",
                                            placeholder=f"Nhập {label.lower()}..."
                                        )
                                        displayed_tags.add(tag)
                                    else:
                                        current_val = st.session_state.get(f"dyn_{tag}", final_dynamic_data.get(tag, ""))
                                        st.text_input(
                                            f"{label} (Đồng bộ)", 
                                            value=current_val, 
                                            key=f"readonly_{tpl}_{tag}",
                                            disabled=True,
                                            help="Đã đồng bộ tự động từ biểu mẫu bên trên."
                                        )
            
            # 2. BẢNG KÊ HÀNG HÓA / DỊCH VỤ
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**2. Bảng kê hàng hóa / dịch vụ:**")
            
            ds = data.get("danh_sach_hang_hoa", [])
            if not ds: 
                ds = [{"ten_hang_hoa": "", "don_vi_tinh": "", "so_luong": 1, "don_gia": 0, "thanh_tien": 0}]
            
            fmt_data = []
            for item in ds:
                sl = float(item.get("so_luong") or 0)
                gia = float(item.get("don_gia") or 0)
                tt = float(item.get("thanh_tien") or (sl * gia))
                fmt_data.append({
                    "Tên hàng hóa": str(item.get("ten_hang_hoa") or ""), 
                    "ĐVT": str(item.get("don_vi_tinh") or ""), 
                    "Số lượng": float(sl), 
                    "Đơn giá": float(gia), 
                    "Thành tiền": float(tt)
                })
            
            edited_df = st.data_editor(
                pd.DataFrame(fmt_data), 
                num_rows="dynamic", 
                use_container_width=True, 
                hide_index=True
            )
            
            danh_sach_da_chinh_sua = edited_df.to_dict('records')
            ds_chuan = []
            sum_thanh_tien = 0.0
            
            for item in danh_sach_da_chinh_sua:
                ten_hh = str(item.get("Tên hàng hóa") or "").strip()
                if not ten_hh or ten_hh == "None": 
                    continue 
                
                sl = float(item.get("Số lượng") or 0)
                gia = float(item.get("Đơn giá") or 0)
                tt = float(item.get("Thành tiền") or 0)
                if tt == 0 and sl > 0 and gia > 0:
                    tt = sl * gia
                    
                sum_thanh_tien += tt
                ds_chuan.append({
                    "ten_hang_hoa": ten_hh, 
                    "don_vi_tinh": str(item.get("ĐVT") or ""),
                    "so_luong": sl, 
                    "don_gia": gia, 
                    "thanh_tien": tt, 
                    "muc_dich_su_dung": "", 
                    "ton_kho": "", 
                    "nha_cung_cap": "", 
                    "ghi_chu": "", 
                    "muc_dich": ""
                })

            # 3. THIẾT LẬP THUẾ VAT & TÍNH TOÁN
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**3. Thiết lập thuế GTGT (VAT):**")
            
            vat_info = data.get("thong_tin_vat", {})
            ai_da_bao_gom_vat = vat_info.get("da_bao_gom_vat", False)
            try:
                ai_thue_suat = float(vat_info.get("thue_suat", 8))
            except (ValueError, TypeError):
                ai_thue_suat = 8.0
                
            col_vat1, col_vat2 = st.columns([1.6, 1])
            with col_vat1:
                vat_mode = st.radio(
                    "Trạng thái thuế trên chứng từ / báo giá gốc:",
                    options=["Chưa gồm VAT (Cộng thêm thuế vào tổng thanh toán)", "Đã bao gồm VAT (Bóc tách tiền thuế từ tổng thanh toán)"],
                    index=1 if ai_da_bao_gom_vat else 0,
                    horizontal=False
                )
                is_vat_included = "Đã bao gồm VAT" in vat_mode
            with col_vat2:
                thue_suat = st.number_input(
                    "Thuế suất GTGT (%):", 
                    min_value=0.0, 
                    max_value=100.0, 
                    value=ai_thue_suat, 
                    step=1.0
                )

            # TÍNH TOÁN THEO NGHIỆP VỤ KẾ TOÁN
            if not is_vat_included:
                # Trường hợp: Báo giá CHƯA gồm VAT -> Thành tiền là tiền hàng, cộng thuế vào
                tong_tien_hang = sum_thanh_tien
                tien_thue = tong_tien_hang * (thue_suat / 100.0)
                tong_thanh_toan = tong_tien_hang + tien_thue
            else:
                # Trường hợp: Báo giá ĐÃ gồm VAT -> Thành tiền trên chứng từ là tổng thanh toán, bóc tách ngược tiền hàng
                tong_thanh_toan = sum_thanh_tien
                tong_tien_hang = tong_thanh_toan / (1.0 + (thue_suat / 100.0)) if thue_suat >= 0 else tong_thanh_toan
                tien_thue = tong_thanh_toan - tong_tien_hang

            chu_so_tien = doc_so_tien_vn(tong_thanh_toan)
            
            # Hiển thị thẻ tóm tắt tài chính chuẩn mực
            display_financial_summary(
                tong_tien_hang=tong_tien_hang,
                tien_thue=tien_thue,
                tong_thanh_toan=tong_thanh_toan,
                chu_so_tien=chu_so_tien,
                is_vat_included=is_vat_included,
                thue_suat=thue_suat
            )

            # 4. XUẤT HỒ SƠ & TẢI VỀ
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📥 XUẤT TRỌN BỘ HỒ SƠ (.ZIP)", type="primary", use_container_width=True):
                # Chuẩn bị dữ liệu cho từng dòng hàng hóa
                for item in ds_chuan:
                    if is_vat_included:
                        item["tong_cong"] = f"{item['thanh_tien']:,.0f}"
                    else:
                        tong_cong_mon = item["thanh_tien"] * (1.0 + (thue_suat / 100.0))
                        item["tong_cong"] = f"{tong_cong_mon:,.0f}"
                
                final_data = {
                    **final_dynamic_data, 
                    "tong_tien_hang": tong_tien_hang, 
                    "thue_gtgt": tien_thue, 
                    "tong_thanh_toan": tong_thanh_toan, 
                    "tong_cong": tong_thanh_toan,
                    "so_tien_bang_chu": chu_so_tien, 
                    "danh_sach_hang_hoa": ds_chuan
                }
                
                try:
                    zip_buffer = io.BytesIO()
                    with tempfile.TemporaryDirectory() as tmp_dir:
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for tpl in selected_templates:
                                tpl_path = os.path.join(TEMPLATE_DIR, tpl)
                                out_path = os.path.join(tmp_dir, f"HoSo_{tpl}")
                                
                                if tpl.endswith('.docx'): 
                                    DocumentBuilder.fill_word_template(tpl_path, out_path, final_data)
                                else: 
                                    DocumentBuilder.fill_excel_template(tpl_path, out_path, final_data)
                                
                                zip_file.write(out_path, arcname=f"HoSo_{tpl}")
                    
                    st.success("Đã hoàn tất lập trọn bộ hồ sơ chứng từ!")
                    st.download_button(
                        label="💾 TẢI VỀ BỘ HỒ SƠ (FILE ZIP)", 
                        data=zip_buffer.getvalue(), 
                        file_name="Bo_Ho_So_Chung_Tu.zip", 
                        mime="application/zip", 
                        use_container_width=True
                    )
                    
                    ten_doi_tac = final_dynamic_data.get('ten_cong_ty_ben_b', ai_flat_data.get('ten_cong_ty', ''))
                    display_zalo_message(ten_doi_tac, tong_thanh_toan)

                except Exception as e:
                    st.error(f"Đã xảy ra lỗi trong quá trình xuất hồ sơ: {str(e)}")