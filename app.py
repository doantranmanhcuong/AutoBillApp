import streamlit as st
import os
import time
import pandas as pd
import zipfile
import io
import tempfile
from docx import Document
try:
    import pymupdf as fitz
except ImportError:
    import fitz

from core.config import API_KEY, API_KEYS, TEMPLATE_DIR
from core.utils.helpers import doc_so_tien_vn, exhaustive_extract_tags, safe_float
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
            doc = None
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
            finally:
                if doc is not None:
                    try:
                        doc.close()
                    except Exception:
                        pass
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

# Kiểm tra thư mục biểu mẫu (bỏ qua file tạm/lock file mở bởi Office có tiền tố ~$)
available_templates = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(('.xlsx', '.docx')) and not f.startswith('~$')]

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
    if '_batch_results' in st.session_state:
        del st.session_state['_batch_results']
    if '_last_invoice_results' in st.session_state:
        del st.session_state['_last_invoice_results']

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
    elif len(temp_files_list) <= 4:
        st.caption(f"📑 Đã nạp **{len(temp_files_list)} hóa đơn** (Bấm chọn từng hóa đơn để xem):")
        file_tabs = st.tabs([f"📄 HĐ #{i+1}" for i in range(len(temp_files_list))])
        for idx, item in enumerate(temp_files_list):
            with file_tabs[idx]:
                st.caption(f"📎 **Tệp ({idx + 1}/{len(temp_files_list)}):** `{item['name']}`")
                render_file_preview(item["name"], item["path"], unique_idx=idx)
    else:
        st.caption(f"📑 Đã nạp **{len(temp_files_list)} hóa đơn**:")
        selected_idx = st.selectbox(
            "Chọn hóa đơn xem trước:",
            options=list(range(len(temp_files_list))),
            format_func=lambda i: f"📄 HĐ #{i+1}: {temp_files_list[i]['name']}"
        )
        item = temp_files_list[selected_idx]
        st.caption(f"📎 **Tệp ({selected_idx + 1}/{len(temp_files_list)}):** `{item['name']}`")
        render_file_preview(item["name"], item["path"], unique_idx=selected_idx)

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
            "ghi_chu", "tr"
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
            if '_batch_results' not in st.session_state:
                st.session_state['_batch_results'] = {}

            btn_label = "🔍 Trích xuất thông tin chứng từ" if len(temp_files_list) == 1 else f"🔍 Trích xuất & Gộp {len(temp_files_list)} hóa đơn"
            if st.button(btn_label, type="primary", use_container_width=True):
                invoice_results = []
                progress_bar = st.progress(0.0)
                status_box = st.empty()
                total_files = len(temp_files_list)
                extractor = AIExtractor(API_KEYS)
                called_ai_count = 0

                for idx, item in enumerate(temp_files_list):
                    fname = item["name"]
                    fpath = item["path"]

                    # Nếu file này đã được bóc tách thành công ở lần trước trong đợt tải này -> Tận dụng ngay, không tốn thời gian chạy lại
                    if fname in st.session_state['_batch_results'] and "error" not in st.session_state['_batch_results'][fname]:
                        cached_data = st.session_state['_batch_results'][fname]
                        status_box.caption(f"⚡ HĐ {idx + 1}/{total_files} `{fname}`: Đã hoàn tất trước đó (lấy kết quả ngay).")
                        invoice_results.append({
                            "filename": fname,
                            "data": cached_data,
                            "path": fpath
                        })
                        progress_bar.progress((idx + 1) / total_files)
                        continue

                    # Giãn cách 1.5 giây giữa các lần gọi AI để bảo vệ hạn ngạch RPM
                    if called_ai_count > 0:
                        status_box.caption(f"⏳ Tạm dừng 1.5s để bảo toàn hạn ngạch máy chủ AI...")
                        time.sleep(1.5)
                    called_ai_count += 1

                    status_box.info(f"⏳ Đang bóc tách hóa đơn {idx + 1}/{total_files}: **{fname}**...")
                    progress_bar.progress(idx / total_files)
                    
                    extracted = extractor.extract_invoice_data(fpath, expected_tags=all_dynamic_tags)
                    if "error" not in extracted:
                        st.session_state['_batch_results'][fname] = extracted

                    invoice_results.append({
                        "filename": fname,
                        "data": extracted,
                        "path": fpath
                    })

                progress_bar.progress(1.0)
                status_box.empty()
                progress_bar.empty()

                st.session_state['_last_invoice_results'] = invoice_results

                if len(invoice_results) == 1:
                    merged_res = invoice_results[0]["data"]
                else:
                    merged_res = InvoiceMerger.merge_invoices(invoice_results)

                if "error" not in merged_res:
                    st.session_state['data'] = merged_res
                    # Xóa các state nhập liệu dyn_ cũ để nạp mới cho các hóa đơn vừa bóc tách
                    for k in list(st.session_state.keys()):
                        if k.startswith("dyn_") or k.startswith("readonly_"):
                            del st.session_state[k]

                    meta = merged_res.get("_multi_invoice_meta", {})
                    failed_cnt = meta.get("failed_count", 0)
                    if failed_cnt > 0:
                        st.warning(f"⚠️ Đã gộp {meta.get('total_invoices')}/{meta.get('total_uploaded')} hóa đơn. Có {failed_cnt} hóa đơn bị lỗi: {', '.join(meta.get('failed_files', []))}")
                    else:
                        success_msg = "Trích xuất thông tin chứng từ thành công!" if len(temp_files_list) == 1 else f"Đã trích xuất & gộp trọn vẹn cả {len(temp_files_list)} hóa đơn thành công!"
                        st.toast(success_msg, icon="✅")
                else:
                    st.error(merged_res['error'])

            # Cứu hộ riêng từng file lỗi để người dùng không phải bóc tách lại cả lô
            last_results = st.session_state.get('_last_invoice_results', [])
            failed_items = [item for item in last_results if "error" in item.get("data", {})]
            if failed_items:
                st.warning(f"⚠️ Có {len(failed_items)} hóa đơn chưa hoàn thành. Các hóa đơn khác đã được giữ nguyên kết quả:")
                for f_idx, fail_item in enumerate(failed_items):
                    f_col1, f_col2 = st.columns([3.5, 1.5])
                    with f_col1:
                        st.caption(f"📄 **{fail_item['filename']}**: {fail_item.get('data', {}).get('error')}")
                    with f_col2:
                        if st.button(f"🔄 Thử lại `{fail_item['filename']}`", key=f"retry_single_{f_idx}", use_container_width=True):
                            with st.spinner(f"Đang bóc tách lại riêng '{fail_item['filename']}'..."):
                                retry_ext = AIExtractor(API_KEYS)
                                retried_data = retry_ext.extract_invoice_data(fail_item['path'], expected_tags=all_dynamic_tags)
                                if "error" not in retried_data:
                                    st.session_state['_batch_results'][fail_item['filename']] = retried_data
                                    for res in last_results:
                                        if res['filename'] == fail_item['filename']:
                                            res['data'] = retried_data
                                    st.session_state['_last_invoice_results'] = last_results

                                    if len(last_results) == 1:
                                        new_merged = last_results[0]["data"]
                                    else:
                                        new_merged = InvoiceMerger.merge_invoices(last_results)

                                    st.session_state['data'] = new_merged
                                    for k in list(st.session_state.keys()):
                                        if k.startswith("dyn_") or k.startswith("readonly_"):
                                            del st.session_state[k]

                                    st.success(f"Bóc tách thành công `{fail_item['filename']}`!")
                                    time.sleep(0.8)
                                    st.rerun()
                                else:
                                    st.error(f"Thử lại thất bại: {retried_data.get('error')}")

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

            # Đồng bộ thông tin nhà cung cấp từ dữ liệu bóc tách hóa đơn
            ncc_extracted = (
                str(ai_flat_data.get("nha_cung_cap") or "") or
                str(ai_flat_data.get("ten_cong_ty") or "") or
                str(ai_flat_data.get("ten_cong_ty_ben_b") or "") or
                str(data.get("thong_tin_nha_cung_cap", {}).get("ten_cong_ty") or "")
            ).strip()

            if ncc_extracted and not ai_flat_data.get("nha_cung_cap"):
                ai_flat_data["nha_cung_cap"] = ncc_extracted

            # Số phiếu đề xuất để người dùng tự nhập (không điền sẵn số hóa đơn vào)
            ai_flat_data["so_phieu"] = ""
            
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
                    "nha_cung_cap": "Nhà cung cấp",
                    "muc_dich": "Mục đích sử dụng",
                    "muc_dich_su_dung": "Mục đích sử dụng",
                    "ton_kho": "Tồn kho",
                    "ton_kh": "Tồn kho",
                    "chuc_vu": "Chức vụ",
                    "chuc_vu_ben_b": "Chức vụ Bên B",
                    "ten_cong_ty_ben_b": "Tên công ty Bên B",
                    "nguoi_dai_dien_ben_b": "Người đại diện Bên B",
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
                                    session_key = f"dyn_{tag}"
                                    if tag not in displayed_tags:
                                        default_val = str(ai_flat_data.get(tag, ""))
                                        if session_key not in st.session_state:
                                            st.session_state[session_key] = default_val
                                        
                                        val_typed = st.text_input(
                                            label, 
                                            key=session_key,
                                            placeholder=f"Nhập {label.lower()}..."
                                        )
                                        final_dynamic_data[tag] = val_typed
                                        displayed_tags.add(tag)
                                    else:
                                        current_val = st.session_state.get(session_key, "")
                                        st.text_input(
                                            f"{label} (Đồng bộ)", 
                                            value=current_val, 
                                            key=f"readonly_{tpl}_{tag}",
                                            disabled=True,
                                            help="Đã đồng bộ tự động từ biểu mẫu bên trên."
                                        )
            
            # 2. BẢNG KÊ HÀNG HÓA / DỊCH VỤ (KHÓA CHẾ ĐỘ CHỈ ĐỌC)
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**2. Bảng kê hàng hóa / dịch vụ:**")
            
            ds = data.get("danh_sach_hang_hoa", [])
            if not ds: 
                ds = [{"ten_hang_hoa": "", "don_vi_tinh": "", "so_luong": 1, "don_gia": 0, "thanh_tien": 0}]
            
            has_ghi_chu = any(bool(str(item.get("ghi_chu") or "").strip()) for item in ds if isinstance(item, dict))
            fmt_data = []
            for idx_h, item in enumerate(ds):
                if not isinstance(item, dict):
                    continue
                sl = safe_float(item.get("so_luong"), 0.0)
                gia = safe_float(item.get("don_gia"), 0.0)
                tt = safe_float(item.get("thanh_tien"), (sl * gia))
                row_dict = {
                    "STT": idx_h + 1,
                    "Tên hàng hóa": str(item.get("ten_hang_hoa") or ""), 
                    "ĐVT": str(item.get("don_vi_tinh") or ""), 
                    "Số lượng": sl, 
                    "Đơn giá": gia, 
                    "Thành tiền": tt
                }
                if has_ghi_chu:
                    row_dict["Ghi chú / Nguồn HĐ"] = str(item.get("ghi_chu") or "").strip()
                fmt_data.append(row_dict)
            
            st.dataframe(
                pd.DataFrame(fmt_data), 
                use_container_width=True, 
                hide_index=True
            )
            st.caption("🔒 **Bảng kê được khóa cố định (Chỉ đọc):** Dữ liệu bảng kê và đơn giá được trích xuất trực tiếp từ hóa đơn/báo giá gốc, không cho phép chỉnh sửa nhằm chống sai lệch và bảo đảm tính minh bạch.")

            # Lấy thông tin chung đồng bộ để gán vào từng dòng bảng kê
            resolved_ncc = (
                str(final_dynamic_data.get("nha_cung_cap") or "") or
                str(final_dynamic_data.get("ten_cong_ty") or "") or
                str(final_dynamic_data.get("ten_cong_ty_ben_b") or "") or
                str(ai_flat_data.get("nha_cung_cap") or "") or
                str(ai_flat_data.get("ten_cong_ty") or "") or
                ncc_extracted
            ).strip()

            resolved_muc_dich = (
                str(final_dynamic_data.get("muc_dich") or "") or
                str(final_dynamic_data.get("muc_dich_su_dung") or "") or
                str(ai_flat_data.get("muc_dich") or "") or
                str(ai_flat_data.get("muc_dich_su_dung") or "")
            ).strip()

            resolved_ton_kho = (
                str(final_dynamic_data.get("ton_kho") or "") or
                str(final_dynamic_data.get("ton_kh") or "") or
                str(ai_flat_data.get("ton_kho") or "")
            ).strip()

            ds_chuan = []
            sum_thanh_tien = 0.0

            for item in ds:
                if not isinstance(item, dict):
                    continue
                ten_hh = str(item.get("ten_hang_hoa") or "").strip()
                if not ten_hh or ten_hh == "None": 
                    continue 
                
                sl = safe_float(item.get("so_luong"), 0.0)
                gia = safe_float(item.get("don_gia"), 0.0)
                tt = safe_float(item.get("thanh_tien"), (sl * gia))
                if tt == 0 and sl > 0 and gia > 0:
                    tt = sl * gia
                    
                sum_thanh_tien += tt

                item_ncc = str(item.get("nha_cung_cap") or "").strip() or resolved_ncc
                item_md = str(item.get("muc_dich") or item.get("muc_dich_su_dung") or "").strip() or resolved_muc_dich
                item_tk = str(item.get("ton_kho") or item.get("ton_kh") or "").strip() or resolved_ton_kho
                item_gc = str(item.get("ghi_chu") or "").strip()

                ds_chuan.append({
                    "stt": len(ds_chuan) + 1,
                    "ten_hang_hoa": ten_hh, 
                    "don_vi_tinh": str(item.get("don_vi_tinh") or ""),
                    "so_luong": sl, 
                    "don_gia": gia, 
                    "thanh_tien": tt, 
                    "muc_dich": item_md, 
                    "muc_dich_su_dung": item_md, 
                    "ton_kho": item_tk, 
                    "ton_kh": item_tk, 
                    "nha_cung_cap": item_ncc, 
                    "ghi_chu": item_gc
                })

            # 3. THIẾT LẬP THUẾ VAT & TÍNH TOÁN (HỖ TRỢ ĐA THUẾ SUẤT 5%, 8%, 10%)
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**3. Thiết lập thuế GTGT (VAT):**")
            
            vat_info = data.get("thong_tin_vat", {}) if isinstance(data, dict) else {}
            ai_da_bao_gom_vat = vat_info.get("da_bao_gom_vat", False) if isinstance(vat_info, dict) else False
            ai_thue_suat = safe_float(vat_info.get("thue_suat", 8), 8.0) if isinstance(vat_info, dict) else 8.0

            # Quét tất cả các mức thuế suất thực tế trong danh sách hàng hóa
            distinct_rates = sorted(list(set(
                safe_float(item.get("thue_suat", ai_thue_suat) if item.get("thue_suat") is not None else ai_thue_suat, ai_thue_suat)
                for item in ds if isinstance(item, dict)
            )))
            is_multi_tax = len(distinct_rates) > 1

            col_vat1, col_vat2 = st.columns([1.4, 1.2])
            with col_vat1:
                vat_mode = st.radio(
                    "Trạng thái thuế trên chứng từ / báo giá gốc:",
                    options=["Chưa gồm VAT (Cộng thêm thuế vào tổng thanh toán)", "Đã bao gồm VAT (Bóc tách tiền thuế từ tổng thanh toán)"],
                    index=1 if ai_da_bao_gom_vat else 0,
                    horizontal=False
                )
                is_vat_included = "Đã bao gồm VAT" in vat_mode

            with col_vat2:
                if is_multi_tax:
                    st.info(f"💡 Phát hiện đơn hàng có **{len(distinct_rates)} mức thuế**: {', '.join([f'{r:g}%' for r in distinct_rates])}")
                    tax_mode = st.radio(
                        "Phương thức tính thuế GTGT:",
                        options=[f"Theo từng mặt hàng ({', '.join([f'{r:g}%' for r in distinct_rates])})", "Áp dụng một mức thuế chung cho tất cả"],
                        index=0,
                        horizontal=False
                    )
                    use_item_tax = "Theo từng mặt hàng" in tax_mode
                    if not use_item_tax:
                        thue_suat_override = st.number_input("Thuế suất chung (%):", min_value=0.0, max_value=100.0, value=ai_thue_suat, step=1.0)
                    else:
                        thue_suat_override = None
                else:
                    use_item_tax = False
                    thue_suat_override = st.number_input(
                        "Thuế suất GTGT (%):", 
                        min_value=0.0, 
                        max_value=100.0, 
                        value=distinct_rates[0] if distinct_rates else ai_thue_suat, 
                        step=1.0
                    )

            # TÍNH TOÁN THEO TỪNG NHÓM THUẾ SUẤT
            tax_breakdown = {}
            for item in ds:
                if not isinstance(item, dict):
                    continue
                sl = safe_float(item.get("so_luong"), 0.0)
                gia = safe_float(item.get("don_gia"), 0.0)
                tt = safe_float(item.get("thanh_tien"), (sl * gia))
                if tt == 0 and sl > 0 and gia > 0:
                    tt = sl * gia

                if use_item_tax:
                    r = safe_float(item.get("thue_suat", ai_thue_suat) if item.get("thue_suat") is not None else ai_thue_suat, ai_thue_suat)
                else:
                    r = safe_float(thue_suat_override if thue_suat_override is not None else ai_thue_suat, ai_thue_suat)

                if is_vat_included:
                    thanh_toan_mon = tt
                    tien_hang_mon = tt / (1.0 + (r / 100.0)) if r >= 0 else tt
                    tien_thue_mon = thanh_toan_mon - tien_hang_mon
                else:
                    tien_hang_mon = tt
                    tien_thue_mon = tien_hang_mon * (r / 100.0)
                    thanh_toan_mon = tien_hang_mon + tien_thue_mon

                if r not in tax_breakdown:
                    tax_breakdown[r] = {"tien_hang": 0.0, "tien_thue": 0.0, "tong_thanh_toan": 0.0, "so_mon": 0}
                tax_breakdown[r]["tien_hang"] += tien_hang_mon
                tax_breakdown[r]["tien_thue"] += tien_thue_mon
                tax_breakdown[r]["tong_thanh_toan"] += thanh_toan_mon
                tax_breakdown[r]["so_mon"] += 1

            tong_tien_hang = sum(v["tien_hang"] for v in tax_breakdown.values())
            tien_thue = sum(v["tien_thue"] for v in tax_breakdown.values())
            tong_thanh_toan = sum(v["tong_thanh_toan"] for v in tax_breakdown.values())
            chu_so_tien = doc_so_tien_vn(tong_thanh_toan)

            # Hiển thị bảng phân rã chi tiết nếu có từ 2 mức thuế khác nhau
            if len(tax_breakdown) > 1:
                st.markdown("##### 📊 Bảng phân rã số tiền thuế theo từng mức thuế:")
                breakdown_table = []
                for r in sorted(tax_breakdown.keys()):
                    bv = tax_breakdown[r]
                    breakdown_table.append({
                        "Nhóm thuế suất": f"Hàng hóa chịu thuế {r:g}%",
                        "Số mặt hàng": f"{bv['so_mon']} mục",
                        "Tiền hàng (chưa thuế)": f"{bv['tien_hang']:,.0f} đ",
                        "Tiền thuế GTGT": f"{bv['tien_thue']:,.0f} đ",
                        "Tổng thanh toán": f"{bv['tong_thanh_toan']:,.0f} đ"
                    })
                breakdown_table.append({
                    "Nhóm thuế suất": "👉 TỔNG CỘNG TOÀN BỘ",
                    "Số mặt hàng": f"{len(ds)} mục",
                    "Tiền hàng (chưa thuế)": f"{tong_tien_hang:,.0f} đ",
                    "Tiền thuế GTGT": f"{tien_thue:,.0f} đ",
                    "Tổng thanh toán": f"{tong_thanh_toan:,.0f} đ"
                })
                st.dataframe(pd.DataFrame(breakdown_table), use_container_width=True, hide_index=True)

            breakdown_str_parts = [f"{r:g}%: {tax_breakdown[r]['tien_thue']:,.0f} đ" for r in sorted(tax_breakdown.keys())]
            tax_summary_detail = " | ".join(breakdown_str_parts) if len(tax_breakdown) > 1 else None

            # Hiển thị thẻ tóm tắt tài chính chuẩn mực
            display_financial_summary(
                tong_tien_hang=tong_tien_hang,
                tien_thue=tien_thue,
                tong_thanh_toan=tong_thanh_toan,
                chu_so_tien=chu_so_tien,
                is_vat_included=is_vat_included,
                thue_suat=distinct_rates[0] if (len(distinct_rates) == 1 and not use_item_tax) else (thue_suat_override if not use_item_tax else None),
                tax_breakdown_str=tax_summary_detail
            )

            # 4. XUẤT HỒ SƠ & TẢI VỀ
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📥 XUẤT TRỌN BỘ HỒ SƠ (.ZIP)", type="primary", use_container_width=True):
                if not selected_templates:
                    st.warning("⚠️ Vui lòng chọn ít nhất 1 biểu mẫu cần xuất ở ô chọn biểu mẫu bên trên.")
                else:
                    active_single_rate = distinct_rates[0] if (len(distinct_rates) == 1 and not use_item_tax) else (thue_suat_override if not use_item_tax else None)
                    for item in ds_chuan:
                        if not isinstance(item, dict):
                            continue
                        if use_item_tax:
                            r = safe_float(item.get("thue_suat", ai_thue_suat) or ai_thue_suat, ai_thue_suat)
                        else:
                            r = safe_float(thue_suat_override if thue_suat_override is not None else ai_thue_suat, ai_thue_suat)

                        item["thue_suat"] = f"{r:g}%"
                        item["tien_thue"] = item["thanh_tien"] * (r / 100.0)
                        if not item.get("gia_niem_yet"):
                            item["gia_niem_yet"] = item.get("don_gia", 0)

                        if is_vat_included:
                            item["tong_cong"] = f"{item['thanh_tien']:,.0f}"
                        else:
                            tong_cong_mon = item["thanh_tien"] * (1.0 + (r / 100.0))
                            item["tong_cong"] = f"{tong_cong_mon:,.0f}"
                    
                    thue_5_val = tax_breakdown.get(5.0, {}).get("tien_thue", 0.0)
                    thue_8_val = tax_breakdown.get(8.0, {}).get("tien_thue", 0.0)
                    thue_10_val = tax_breakdown.get(10.0, {}).get("tien_thue", 0.0)
                    hang_5_val = tax_breakdown.get(5.0, {}).get("tien_hang", 0.0)
                    hang_8_val = tax_breakdown.get(8.0, {}).get("tien_hang", 0.0)
                    hang_10_val = tax_breakdown.get(10.0, {}).get("tien_hang", 0.0)

                    final_data = {
                        **final_dynamic_data, 
                        "so_phieu": final_dynamic_data.get("so_phieu", ""),
                        "nha_cung_cap": resolved_ncc,
                        "muc_dich": resolved_muc_dich,
                        "muc_dich_su_dung": resolved_muc_dich,
                        "ton_kho": resolved_ton_kho,
                        "ton_kh": resolved_ton_kho,
                        "tong_tien_hang": tong_tien_hang, 
                        "thue_gtgt": tien_thue, 
                        "tong_thanh_toan": tong_thanh_toan, 
                        "tong_cong": tong_thanh_toan,
                        "so_tien_bang_chu": chu_so_tien, 
                        "thue_suat": active_single_rate,
                        "tax_breakdown": tax_breakdown,
                        "thue_gtgt_5": thue_5_val,
                        "thue_5": thue_5_val,
                        "thue_gtgt_8": thue_8_val,
                        "thue_8": thue_8_val,
                        "thue_gtgt_10": thue_10_val,
                        "thue_10": thue_10_val,
                        "tien_hang_5": hang_5_val,
                        "tien_hang_8": hang_8_val,
                        "tien_hang_10": hang_10_val,
                        "chi_tiet_thue": tax_summary_detail or "",
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