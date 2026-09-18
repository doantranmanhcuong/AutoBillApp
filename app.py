import streamlit as st
import os
import re
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

@st.cache_data(show_spinner=False)
def get_pdf_preview_images(file_path: str, dpi: int = 140) -> list:
    """Chuyển đổi các trang PDF thành ảnh PNG bytes và lưu bộ nhớ đệm để tránh render lại gây mờ/giật màn hình khi điền form."""
    images = []
    doc = None
    try:
        doc = fitz.open(file_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=dpi)
            images.append(pix.tobytes("png"))
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass
    return images

def render_file_preview(file_name: str, file_path: str, unique_idx: int = 0):
    """Hiển thị nội dung xem trước tài liệu gốc (PDF, Ảnh, Excel, Word)"""
    file_ext = file_name.split('.')[-1].lower()
    with st.container(border=True):
        if file_ext in ['png', 'jpg', 'jpeg']: 
            st.image(file_path, use_container_width=True)
        elif file_ext == 'pdf':
            try:
                images = get_pdf_preview_images(file_path)
                st.caption(f"📄 Tài liệu PDF gồm {len(images)} trang:")
                for page_num, img_bytes in enumerate(images):
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
    if 'ready_zip' in st.session_state:
        del st.session_state['ready_zip']

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
                    if 'ready_zip' in st.session_state:
                        del st.session_state['ready_zip']

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

            # Trải phẳng dữ liệu (Flatten) theo phân cấp chặt chẽ:
            ai_flat_data = {}
            chung_data = data.get("thong_tin_chung") or {}
            kh_data = data.get("thong_tin_khach_hang") or {}
            ncc_data = data.get("thong_tin_nha_cung_cap") or {}
            tc_data = data.get("tong_ket_tien") or {}
            dong_data = data.get("thong_tin_dong") or {}

            # 1. Nạp thông tin chung & thông tin động
            if isinstance(chung_data, dict):
                for k, v in chung_data.items():
                    if v is not None:
                        ai_flat_data[k] = v
            if isinstance(dong_data, dict):
                for k, v in dong_data.items():
                    if v is not None and k not in ai_flat_data:
                        ai_flat_data[k] = v
            if isinstance(tc_data, dict):
                for k, v in tc_data.items():
                    if v is not None:
                        ai_flat_data[k] = v

            # 2. Thông tin khách hàng (Bên Mua / Bên A) - Graphene Global
            # TUYỆT ĐỐI KHÔNG gán thông tin khách hàng vào các key chung (ten_cong_ty, so_tai_khoan)
            kh_mst = str(kh_data.get("ma_so_thue") or "").strip()
            kh_stk = str(kh_data.get("so_tai_khoan") or "").strip()
            kh_ten = str(kh_data.get("ten_khach_hang") or kh_data.get("ten_cong_ty") or "").strip()
            if kh_ten:
                ai_flat_data["ten_cong_ty_ben_a"] = kh_ten
            if kh_mst:
                ai_flat_data["mst_ben_a"] = kh_mst
            if kh_stk:
                ai_flat_data["so_tai_khoan_ben_a"] = kh_stk
            if kh_data.get("dia_chi"):
                ai_flat_data["dia_chi_ben_a"] = str(kh_data.get("dia_chi")).strip()
            if kh_data.get("ten_ngan_hang"):
                ai_flat_data["ten_ngan_hang_ben_a"] = str(kh_data.get("ten_ngan_hang")).strip()
            if kh_data.get("nguoi_dai_dien"):
                ai_flat_data["nguoi_dai_dien_ben_a"] = str(kh_data.get("nguoi_dai_dien")).strip()
            if kh_data.get("chuc_vu"):
                ai_flat_data["chuc_vu_ben_a"] = str(kh_data.get("chuc_vu")).strip()

            # 3. THÔNG TIN NHÀ CUNG CẤP (BÊN BÁN / ĐƠN VỊ THỤ HƯỞNG) BẮT BUỘC ƯU TIÊN CAO NHẤT
            if isinstance(ncc_data, dict):
                ncc_ten = str(ncc_data.get("ten_cong_ty") or "").strip()
                ncc_mst = str(ncc_data.get("ma_so_thue") or "").strip()
                ncc_dc = str(ncc_data.get("dia_chi") or "").strip()
                ncc_dt = str(ncc_data.get("dien_thoai") or "").strip()
                ncc_stk = str(ncc_data.get("so_tai_khoan") or "").strip()
                ncc_nh = str(ncc_data.get("ten_ngan_hang") or "").strip()
                ncc_dd = str(ncc_data.get("nguoi_dai_dien") or ncc_data.get("dai_dien") or "").strip()
                ncc_cv = str(ncc_data.get("chuc_vu") or "").strip()

                # BIỆN PHÁP CHỐNG BỊA ĐẶT & CHỐNG NHẦM THÔNG TIN BÊN MUA (GRAPHENE) SANG BÊN BÁN (NCC):
                if "graphene" in ncc_ten.lower():
                    ncc_ten = ""
                if ncc_mst in ("4101649609", kh_mst):
                    ncc_mst = ""
                if ncc_stk in ("884249867", kh_stk):
                    ncc_stk = ""

                if ncc_ten:
                    ai_flat_data["ten_cong_ty"] = ncc_ten
                    ai_flat_data["nha_cung_cap"] = ncc_ten
                    ai_flat_data["ten_cong_ty_ben_b"] = ncc_ten
                if ncc_mst:
                    ai_flat_data["ma_so_thue"] = ncc_mst
                    ai_flat_data["mst_ben_b"] = ncc_mst
                if ncc_dc:
                    ai_flat_data["dia_chi"] = ncc_dc
                    ai_flat_data["dia_chi_ben_b"] = ncc_dc
                if ncc_dt:
                    ai_flat_data["dien_thoai"] = ncc_dt
                    ai_flat_data["dien_thoai_ben_b"] = ncc_dt
                if ncc_stk:
                    ai_flat_data["so_tai_khoan"] = ncc_stk
                    ai_flat_data["so_tai_khoan_ben_b"] = ncc_stk
                if ncc_nh:
                    ai_flat_data["ten_ngan_hang"] = ncc_nh
                if ncc_dd:
                    ai_flat_data["nguoi_dai_dien_ben_b"] = ncc_dd
                if ncc_cv:
                    ai_flat_data["chuc_vu_ben_b"] = ncc_cv

            # Quét dọn ai_flat_data chống rò rỉ thông tin Bên Mua (Graphene)
            if str(ai_flat_data.get("ma_so_thue") or "").strip() in ("4101649609", kh_mst):
                ai_flat_data["ma_so_thue"] = ""
                ai_flat_data["mst_ben_b"] = ""
            if str(ai_flat_data.get("so_tai_khoan") or "").strip() in ("884249867", kh_stk):
                ai_flat_data["so_tai_khoan"] = ""
                ai_flat_data["so_tai_khoan_ben_b"] = ""
            if "graphene" in str(ai_flat_data.get("ten_cong_ty") or "").lower():
                ai_flat_data["ten_cong_ty"] = ""
                ai_flat_data["nha_cung_cap"] = ""
                ai_flat_data["ten_cong_ty_ben_b"] = ""

            # Đồng bộ tên nhà cung cấp
            ncc_extracted = (
                str(ai_flat_data.get("nha_cung_cap") or "") or
                str(ai_flat_data.get("ten_cong_ty") or "") or
                str(ai_flat_data.get("ten_cong_ty_ben_b") or "") or
                str(ncc_data.get("ten_cong_ty") or "")
            ).strip()

            if "graphene" in ncc_extracted.lower():
                ncc_extracted = ""

            if ncc_extracted and not ai_flat_data.get("nha_cung_cap"):
                ai_flat_data["nha_cung_cap"] = ncc_extracted

            # Đảm bảo các trường số phiếu, số hợp đồng, số hóa đơn không bị tự ý bịa đặt
            ai_flat_data["so_phieu"] = ""
            if not ai_flat_data.get("so_hop_dong"):
                ai_flat_data["so_hop_dong"] = ""
            if not ai_flat_data.get("so_chung_tu"):
                ai_flat_data["so_chung_tu"] = ""

            # Tự động trích xuất ngày, tháng, năm cho các hợp đồng nếu có ngày tháng chứng từ
            raw_date = str(ai_flat_data.get("ngay_thang_nam") or ai_flat_data.get("ngay_lap_phieu") or "").strip()
            if raw_date:
                m_d = re.search(r'ngày\s*(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})', raw_date, re.I)
                if not m_d:
                    m_d = re.search(r'(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})', raw_date)
                if not m_d:
                    m_d = re.search(r'(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})', raw_date)
                    if m_d:
                        ai_flat_data.setdefault("nam_ky", m_d.group(1))
                        ai_flat_data.setdefault("thang_ky", m_d.group(2).zfill(2))
                        ai_flat_data.setdefault("ngay_ky", m_d.group(3).zfill(2))
                if m_d and "nam_ky" not in ai_flat_data:
                    ai_flat_data.setdefault("ngay_ky", m_d.group(1).zfill(2))
                    ai_flat_data.setdefault("thang_ky", m_d.group(2).zfill(2))
                    ai_flat_data.setdefault("nam_ky", m_d.group(3))
            
            final_dynamic_data = {} 
            
            # 1. KIỂM DUYỆT VÀ NHẬP BỔ SUNG THÔNG TIN CÁC BIỂU MẪU
            if template_tags_map:
                st.markdown("**1. Kiểm duyệt & Nhập bổ sung thông tin theo biểu mẫu đã chọn:**")
                
                TAG_MAPPING = {
                    "ten_cong_ty": "Tên NCC / Đơn vị thụ hưởng (Bên B)",
                    "nha_cung_cap": "Tên NCC / Đơn vị thụ hưởng (Bên B)",
                    "ten_cong_ty_ben_b": "Tên NCC / Đơn vị thụ hưởng (Bên B)",
                    "so_tai_khoan": "Số tài khoản ngân hàng NCC",
                    "so_tai_khoan_ben_b": "Số tài khoản ngân hàng NCC (Bên B)",
                    "ten_ngan_hang": "Tên ngân hàng NCC / Đơn vị thụ hưởng",
                    "ma_so_thue": "Mã số thuế NCC (Bên B)",
                    "mst_ben_b": "Mã số thuế NCC (Bên B)",
                    "dia_chi": "Địa chỉ NCC (Bên B)",
                    "dia_chi_ben_b": "Địa chỉ NCC (Bên B)",
                    "dien_thoai": "Số điện thoại NCC",
                    "dien_thoai_ben_b": "Số điện thoại NCC (Bên B)",
                    "nguoi_dai_dien_ben_b": "Người đại diện Bên B (NCC)",
                    "chuc_vu_ben_b": "Chức vụ người đại diện Bên B",

                    "ho_ten_nguoi_de_nghi": "Họ tên người đề nghị / đề xuất",
                    "nguoi_de_xuat": "Họ tên người đề nghị / đề xuất",
                    "bo_phan": "Bộ phận / Phòng ban",
                    "bo_phan_cong_tac": "Bộ phận công tác",
                    "chuc_vu": "Chức vụ người đề nghị",
                    "ly_do_de_nghi": "Lý do đề nghị / thanh toán",
                    "ly_do_thanh_toan": "Lý do thanh toán",
                    "hinh_thuc_thanh_toan": "Hình thức thanh toán",
                    "thoi_han_hoan_ung": "Thời hạn hoàn ứng",
                    "so_phieu": "Số phiếu đề xuất",
                    "muc_dich": "Mục đích sử dụng",
                    "muc_dich_su_dung": "Mục đích sử dụng",
                    "ton_kho": "Tình trạng tồn kho",
                    "ton_kh": "Tình trạng tồn kho",
                    "so_don_dat_hang": "Số đơn đặt hàng",
                    "so_de_xuat": "Số đề xuất",
                    "nguoi_phu_trach": "Người phụ trách mua hàng",
                    "email_phu_trach": "Email người phụ trách",
                    "nguoi_nhan_hang": "Người nhận hàng",
                    "sdt_nguoi_nhan": "SĐT người nhận hàng",

                    "so_hop_dong": "Số hợp đồng kinh tế",
                    "so_chung_tu": "Số hóa đơn GTGT / Chứng từ gốc",
                    "so_hoa_don": "Số hóa đơn GTGT / Chứng từ gốc",
                    "ngay_ky": "Ngày ký",
                    "thang_ky": "Tháng ký",
                    "nam_ky": "Năm ký",
                    "ngay_lap_phieu": "Ngày lập phiếu",
                    "ngay_thang_nam": "Ngày tháng năm lập phiếu",
                    "thoi_gian_thuc_hien": "Thời gian thực hiện hợp đồng/dịch vụ",
                    "dia_diem_thuc_hien": "Địa điểm thực hiện hợp đồng/dịch vụ",
                    "thoi_gian_bao_hanh": "Thời gian bảo hành",
                    "ty_le_tam_ung": "Tỷ lệ tạm ứng (%)"
                }
                
                TAG_ALIASES = {
                    "nha_cung_cap": "ten_cong_ty",
                    "ten_cong_ty_ben_b": "ten_cong_ty",
                    "ten_nha_cung_cap": "ten_cong_ty",
                    "so_tai_khoan_ben_b": "so_tai_khoan",
                    "tai_khoan_ngan_hang": "so_tai_khoan",
                    "stk_ncc": "so_tai_khoan",
                    "dia_chi_ben_b": "dia_chi",
                    "dia_chi_ncc": "dia_chi",
                    "mst_ben_b": "ma_so_thue",
                    "dien_thoai_ben_b": "dien_thoai",
                    "ngan_hang": "ten_ngan_hang",
                    "muc_dich_su_dung": "muc_dich",
                    "ton_kh": "ton_kho",
                    "nguoi_de_xuat": "ho_ten_nguoi_de_nghi",
                    "bo_phan_cong_tac": "bo_phan",
                    "ly_do_thanh_toan": "ly_do_de_nghi",
                    "ngay_lap_phieu": "ngay_thang_nam",
                    "so_hoa_don": "so_chung_tu",
                    "so_tien_viet_bang_chu": "so_tien_bang_chu",
                    "tong_tien_thanh_toan": "tong_thanh_toan",
                    "tong_tien_truoc_thue": "tong_tien_hang"
                }

                CATEGORIES = [
                    {
                        "title": "THÔNG TIN NHÀ CUNG CẤP & ĐƠN VỊ THỤ HƯỞNG (BÊN B)",
                        "icon": "🏢",
                        "description": "Thông tin pháp lý, mã số thuế và tài khoản ngân hàng của đơn vị nhận thanh toán.",
                        "keys": [
                            "ten_cong_ty", "so_tai_khoan", "ten_ngan_hang",
                            "ma_so_thue", "dia_chi", "dien_thoai",
                            "nguoi_dai_dien_ben_b", "chuc_vu_ben_b"
                        ]
                    },
                    {
                        "title": "THÔNG TIN NỘI BỘ & ĐỀ NGHỊ THANH TOÁN",
                        "icon": "📝",
                        "description": "Thông tin người đề nghị, bộ phận, lý do thanh toán/tạm ứng và quy trình mua sắm.",
                        "keys": [
                            "ho_ten_nguoi_de_nghi", "bo_phan", "chuc_vu",
                            "ly_do_de_nghi", "hinh_thuc_thanh_toan", "thoi_han_hoan_ung",
                            "so_phieu", "muc_dich", "ton_kho",
                            "so_don_dat_hang", "so_de_xuat", "nguoi_phu_trach",
                            "email_phu_trach", "nguoi_nhan_hang", "sdt_nguoi_nhan"
                        ]
                    },
                    {
                        "title": "THỜI GIAN, HỢP ĐỒNG & CHỨNG TỪ",
                        "icon": "📅",
                        "description": "Số hợp đồng, số hóa đơn chứng từ, ngày tháng và các điều khoản giao hàng.",
                        "keys": [
                            "so_hop_dong", "so_chung_tu", "ngay_ky", "thang_ky", "nam_ky",
                            "ngay_thang_nam", "thoi_gian_thuc_hien", "dia_diem_thuc_hien",
                            "thoi_gian_bao_hanh", "ty_le_tam_ung"
                        ]
                    }
                ]

                # Tập hợp tất cả canonical keys được yêu cầu bởi các biểu mẫu đang chọn
                needed_canonical_map = {}
                for tpl in selected_templates:
                    for tag in template_tags_map.get(tpl, set()):
                        c_key = TAG_ALIASES.get(tag, tag)
                        if c_key not in needed_canonical_map:
                            needed_canonical_map[c_key] = set()
                        needed_canonical_map[c_key].add(tpl)

                # BẮT BUỘC luôn hiển thị Số hợp đồng và Số hóa đơn / Chứng từ trên form để người dùng tự nhập
                if "so_hop_dong" not in needed_canonical_map:
                    needed_canonical_map["so_hop_dong"] = {"Thông tin hợp đồng"}
                if "so_chung_tu" not in needed_canonical_map:
                    needed_canonical_map["so_chung_tu"] = {"Thông tin hóa đơn / chứng từ"}

                # Kiểm tra các trường bị thiếu trên hóa đơn/báo giá gốc
                missing_fields_list = []
                for c_key in needed_canonical_map:
                    val_check = str(st.session_state.get(f"dyn_{c_key}") or ai_flat_data.get(c_key) or "").strip()
                    if not val_check:
                        missing_fields_list.append(c_key)

                # Banner thông báo trạng thái dữ liệu
                if missing_fields_list:
                    st.warning(
                        f"⚠️ **PHÁT HIỆN {len(missing_fields_list)} TRƯỜNG THÔNG TIN KHÔNG CÓ TRÊN HÓA ĐƠN / BÁO GIÁ GỐC:**\n\n"
                        "Hệ thống tuân thủ nguyên tắc **tuyệt đối không bịa đặt hoặc suy diễn** thông tin còn thiếu. "
                        "Vui lòng kiểm tra và nhập bổ sung vào các ô có ký hiệu **⚠️ (Chưa có trên HĐ - Nhập bổ sung)** bên dưới để hồ sơ xuất ra chuẩn xác nhất."
                    )
                else:
                    st.success("✅ **TOÀN BỘ THÔNG TIN CẦN THIẾT ĐÃ ĐƯỢC TRÍCH XUẤT ĐẦY ĐỦ TỪ CHỨNG TỪ!**")

                all_cat_keys = set().union(*[cat["keys"] for cat in CATEGORIES])
                other_needed_keys = [k for k in needed_canonical_map if k not in all_cat_keys]
                if other_needed_keys:
                    CATEGORIES.append({
                        "title": "THÔNG TIN KHÁC THEO BIỂU MẪU",
                        "icon": "📦",
                        "description": "Các trường bổ sung khác theo yêu cầu đặc thù của biểu mẫu.",
                        "keys": other_needed_keys
                    })

                # Hiển thị form theo từng nhóm logic rõ ràng
                for cat in CATEGORIES:
                    cat_keys = [k for k in cat["keys"] if k in needed_canonical_map]
                    if not cat_keys:
                        continue

                    # Đếm số trường đã có giá trị trong nhóm này
                    filled_in_cat = sum(1 for k in cat_keys if str(st.session_state.get(f"dyn_{k}") or ai_flat_data.get(k) or "").strip())
                    total_in_cat = len(cat_keys)

                    expander_label = f"{cat['icon']} {cat['title']}"
                    status_badge = f"({filled_in_cat}/{total_in_cat} trường đã có" + (f" — ⚠️ Còn thiếu {total_in_cat - filled_in_cat} mục)" if filled_in_cat < total_in_cat else " — ✅ Đầy đủ)")

                    with st.expander(expander_label, expanded=True):
                        st.caption(f"{cat['description']} • **{status_badge}**")
                        cols = st.columns(2)
                        for idx, c_key in enumerate(cat_keys):
                            col = cols[idx % 2]
                            session_key = f"dyn_{c_key}"

                            # Lấy giá trị ban đầu nếu chưa có trong session_state
                            if session_key not in st.session_state:
                                default_val = str(ai_flat_data.get(c_key) or "")
                                st.session_state[session_key] = default_val

                            # Xác định trạng thái thiếu thông tin dựa trên dữ liệu trích xuất ban đầu ai_flat_data
                            # Giữ nhãn tĩnh không đổi để không làm mất focus/con trỏ khi người dùng đang gõ
                            had_extracted = bool(str(ai_flat_data.get(c_key) or "").strip())
                            raw_label = TAG_MAPPING.get(c_key, c_key.replace("_", " ").title())

                            if not had_extracted:
                                field_label = f"⚠️ {raw_label} (Chưa có trên HĐ - Nhập bổ sung)"
                                ph_text = f"Vui lòng nhập {raw_label.lower()}..."
                            else:
                                field_label = f"✅ {raw_label}"
                                ph_text = ""

                            tpl_list = sorted(list(needed_canonical_map[c_key]))
                            help_msg = f"Được sử dụng trong: {', '.join(tpl_list)}"

                            with col:
                                typed = st.text_input(
                                    field_label,
                                    key=session_key,
                                    placeholder=ph_text,
                                    help=help_msg
                                )
                                final_dynamic_data[c_key] = typed

                # Đồng bộ toàn diện các bí danh (aliases)
                for alias_k, canon_k in TAG_ALIASES.items():
                    val = final_dynamic_data.get(canon_k) or st.session_state.get(f"dyn_{canon_k}")
                    if val is not None:
                        final_dynamic_data[alias_k] = val
                        final_dynamic_data[canon_k] = val
                    elif alias_k in final_dynamic_data:
                        final_dynamic_data[canon_k] = final_dynamic_data[alias_k]

                # Đồng bộ trực tiếp cho mọi tag có trong các biểu mẫu đã chọn
                for tpl in selected_templates:
                    for tag in template_tags_map.get(tpl, set()):
                        c_key = TAG_ALIASES.get(tag, tag)
                        val = final_dynamic_data.get(c_key) or final_dynamic_data.get(tag) or st.session_state.get(f"dyn_{c_key}")
                        if val is not None:
                            final_dynamic_data[tag] = val
            
            # Lấy thông tin chung đồng bộ để gán vào từng dòng bảng kê
            resolved_ncc = (
                str(final_dynamic_data.get("nha_cung_cap") or "") or
                str(final_dynamic_data.get("ten_cong_ty") or "") or
                str(final_dynamic_data.get("ten_cong_ty_ben_b") or "") or
                str(ai_flat_data.get("nha_cung_cap") or "") or
                str(ai_flat_data.get("ten_cong_ty") or "") or
                ncc_extracted
            ).strip()

            if "graphene" in resolved_ncc.lower():
                resolved_ncc = ""

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

            # 2. THIẾT LẬP THUẾ VAT & TÍNH TOÁN (HỖ TRỢ ĐA THUẾ SUẤT 5%, 8%, 10%)
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**2. Thiết lập thuế GTGT (VAT):**")
            
            ds = data.get("danh_sach_hang_hoa", [])
            if not ds: 
                ds = [{"ten_hang_hoa": "", "don_vi_tinh": "", "so_luong": 1, "don_gia": 0, "thanh_tien": 0}]

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

            # 3. BẢNG KÊ HÀNG HÓA & TÍNH TOÁN ĐỒNG BỘ 100% (KHÔNG XUNG ĐỘT SỐ LIỆU)
            ds_chuan = []
            fmt_data = []
            tax_breakdown = {}
            has_ghi_chu = any(bool(str(item.get("ghi_chu") or "").strip()) for item in ds if isinstance(item, dict))

            for idx_h, item in enumerate(ds):
                if not isinstance(item, dict):
                    continue
                ten_hh = str(item.get("ten_hang_hoa") or "").strip()
                if not ten_hh or ten_hh == "None":
                    continue

                sl = safe_float(item.get("so_luong"), 0.0)
                gia_raw = safe_float(item.get("don_gia"), 0.0)
                tt_raw = safe_float(item.get("thanh_tien"), (sl * gia_raw))
                if tt_raw == 0 and sl > 0 and gia_raw > 0:
                    tt_raw = sl * gia_raw
                elif gia_raw == 0 and sl > 0 and tt_raw > 0:
                    gia_raw = tt_raw / sl

                if use_item_tax:
                    r = safe_float(item.get("thue_suat", ai_thue_suat) if item.get("thue_suat") is not None else ai_thue_suat, ai_thue_suat)
                else:
                    r = safe_float(thue_suat_override if thue_suat_override is not None else ai_thue_suat, ai_thue_suat)

                if is_vat_included:
                    # Đã gồm VAT: Bóc tách tiền thuế từ tổng số tiền trên báo giá
                    # TUYỆT ĐỐI KHÔNG sửa đơn giá và thành tiền trên báo giá
                    thanh_toan_mon = tt_raw
                    tien_thue_mon = thanh_toan_mon - (thanh_toan_mon / (1.0 + (r / 100.0))) if r > 0 else 0.0
                    tien_hang_mon = thanh_toan_mon - tien_thue_mon
                else:
                    # Chưa gồm VAT: Thành tiền trên báo giá là tiền hàng chưa thuế, cộng thêm thuế
                    tien_hang_mon = tt_raw
                    tien_thue_mon = tien_hang_mon * (r / 100.0)
                    thanh_toan_mon = tien_hang_mon + tien_thue_mon

                if r not in tax_breakdown:
                    tax_breakdown[r] = {"tien_hang": 0.0, "tien_thue": 0.0, "tong_thanh_toan": 0.0, "so_mon": 0}
                tax_breakdown[r]["tien_hang"] += tien_hang_mon
                tax_breakdown[r]["tien_thue"] += tien_thue_mon
                tax_breakdown[r]["tong_thanh_toan"] += thanh_toan_mon
                tax_breakdown[r]["so_mon"] += 1

                item_ncc = str(item.get("nha_cung_cap") or "").strip() or resolved_ncc
                item_md = str(item.get("muc_dich") or item.get("muc_dich_su_dung") or "").strip() or resolved_muc_dich
                item_tk = str(item.get("ton_kho") or item.get("ton_kh") or "").strip() or resolved_ton_kho
                item_gc = str(item.get("ghi_chu") or "").strip()
                item_gny = item.get("gia_niem_yet") or gia_raw
                item_gm = item.get("gia_mua") or gia_raw

                # Dòng chuẩn bị cho xuất file biểu mẫu:
                # BẢO ĐẢM KHỚP 100% ĐƠN GIÁ VÀ THÀNH TIỀN TRÊN BÁO GIÁ GỐC (TUYỆT ĐỐI KHÔNG SỬA)
                ds_chuan.append({
                    "stt": len(ds_chuan) + 1,
                    "ten_hang_hoa": ten_hh,
                    "don_vi_tinh": str(item.get("don_vi_tinh") or ""),
                    "so_luong": sl,
                    "don_gia": gia_raw,          # Giữ nguyên 100% đơn giá từ file báo giá
                    "thanh_tien": tt_raw,        # Giữ nguyên 100% thành tiền từ file báo giá
                    "thue_suat": f"{r:g}%",
                    "tien_thue": tien_thue_mon,
                    "tong_cong": thanh_toan_mon,
                    "gia_niem_yet": item_gny,
                    "gia_mua": item_gm,
                    "muc_dich": item_md,
                    "muc_dich_su_dung": item_md,
                    "ton_kho": item_tk,
                    "ton_kh": item_tk,
                    "nha_cung_cap": item_ncc,
                    "ghi_chu": item_gc
                })

                # Dòng hiển thị xem trước giao diện: Giữ nguyên 100% Đơn giá và Thành tiền theo file báo giá
                row_dict = {
                    "STT": len(ds_chuan),
                    "Tên hàng hóa": ten_hh,
                    "ĐVT": str(item.get("don_vi_tinh") or ""),
                    "Số lượng": sl,
                    "Đơn giá (theo báo giá)": f"{gia_raw:,.0f}",
                    "Thành tiền (theo báo giá)": f"{tt_raw:,.0f}",
                    "Thuế suất": f"{r:g}%",
                    "Tiền thuế VAT": f"{tien_thue_mon:,.0f}",
                    "Tổng thanh toán": f"{thanh_toan_mon:,.0f}"
                }
                if has_ghi_chu:
                    row_dict["Ghi chú / Nguồn HĐ"] = item_gc
                fmt_data.append(row_dict)

            tong_tien_hang = sum(v["tien_hang"] for v in tax_breakdown.values())
            tien_thue = sum(v["tien_thue"] for v in tax_breakdown.values())
            tong_thanh_toan = sum(v["tong_thanh_toan"] for v in tax_breakdown.values())
            chu_so_tien = doc_so_tien_vn(tong_thanh_toan)

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**3. Bảng kê hàng hóa / dịch vụ:**")
            st.dataframe(pd.DataFrame(fmt_data), use_container_width=True, hide_index=True)
            if is_vat_included:
                st.caption(
                    f"🔒 **Bảng kê khớp 100% theo chứng từ / báo giá gốc:** "
                    f"Tổng thanh toán (theo báo giá): **{tong_thanh_toan:,.0f} đ** | "
                    f"Thuế GTGT: **{tien_thue:,.0f} đ** | "
                    f"Tiền hàng (chưa thuế): **{tong_tien_hang:,.0f} đ**"
                )
            else:
                st.caption(
                    f"🔒 **Bảng kê khớp 100% theo chứng từ / báo giá gốc:** "
                    f"Tiền hàng (theo báo giá): **{tong_tien_hang:,.0f} đ** | "
                    f"Thuế GTGT: **{tien_thue:,.0f} đ** | "
                    f"Tổng thanh toán (gồm VAT): **{tong_thanh_toan:,.0f} đ**"
                )

            # Bảng phân rã chi tiết nếu có từ 2 mức thuế
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
                    "Số mặt hàng": f"{len(ds_chuan)} mục",
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
                    with st.spinner("⏳ Đang tự động điền dữ liệu vào các biểu mẫu và nén file ZIP... Vui lòng đợi trong giây lát."):
                        active_single_rate = distinct_rates[0] if (len(distinct_rates) == 1 and not use_item_tax) else (thue_suat_override if not use_item_tax else None)
                        thue_5_val = tax_breakdown.get(5.0, {}).get("tien_thue", 0.0)
                        thue_8_val = tax_breakdown.get(8.0, {}).get("tien_thue", 0.0)
                        thue_10_val = tax_breakdown.get(10.0, {}).get("tien_thue", 0.0)
                        hang_5_val = tax_breakdown.get(5.0, {}).get("tien_hang", 0.0)
                        hang_8_val = tax_breakdown.get(8.0, {}).get("tien_hang", 0.0)
                        hang_10_val = tax_breakdown.get(10.0, {}).get("tien_hang", 0.0)

                        # Đồng bộ định dạng ngày tháng
                        ngay_k = str(final_dynamic_data.get("ngay_ky") or ai_flat_data.get("ngay_ky") or "").strip()
                        thang_k = str(final_dynamic_data.get("thang_ky") or ai_flat_data.get("thang_ky") or "").strip()
                        nam_k = str(final_dynamic_data.get("nam_ky") or ai_flat_data.get("nam_ky") or "").strip()
                        ngay_tn = str(final_dynamic_data.get("ngay_thang_nam") or ai_flat_data.get("ngay_thang_nam") or "").strip()

                        if ngay_k and thang_k and nam_k:
                            ngay_tn_formatted = f"ngày {ngay_k} tháng {thang_k} năm {nam_k}"
                            if not ngay_tn or (final_dynamic_data.get("ngay_ky") or final_dynamic_data.get("thang_ky") or final_dynamic_data.get("nam_ky")):
                                ngay_tn = ngay_tn_formatted
                        elif not ngay_tn and ngay_k and thang_k and nam_k:
                            ngay_tn = f"ngày {ngay_k} tháng {thang_k} năm {nam_k}"

                        final_data = {
                            **ai_flat_data,
                            **final_dynamic_data, 
                            "so_phieu": final_dynamic_data.get("so_phieu", "") or ai_flat_data.get("so_phieu", ""),
                            "ten_cong_ty": resolved_ncc,
                            "nha_cung_cap": resolved_ncc,
                            "ten_cong_ty_ben_b": resolved_ncc,
                            "so_tai_khoan": final_dynamic_data.get("so_tai_khoan") or ai_flat_data.get("so_tai_khoan", ""),
                            "so_tai_khoan_ben_b": final_dynamic_data.get("so_tai_khoan_ben_b") or final_dynamic_data.get("so_tai_khoan") or ai_flat_data.get("so_tai_khoan", ""),
                            "ten_ngan_hang": final_dynamic_data.get("ten_ngan_hang") or ai_flat_data.get("ten_ngan_hang", ""),
                            "ma_so_thue": final_dynamic_data.get("ma_so_thue") or ai_flat_data.get("ma_so_thue", ""),
                            "mst_ben_b": final_dynamic_data.get("mst_ben_b") or final_dynamic_data.get("ma_so_thue") or ai_flat_data.get("ma_so_thue", ""),
                            "dia_chi": final_dynamic_data.get("dia_chi") or ai_flat_data.get("dia_chi", ""),
                            "dia_chi_ben_b": final_dynamic_data.get("dia_chi_ben_b") or final_dynamic_data.get("dia_chi") or ai_flat_data.get("dia_chi", ""),
                            "dien_thoai": final_dynamic_data.get("dien_thoai") or ai_flat_data.get("dien_thoai", ""),
                            "dien_thoai_ben_b": final_dynamic_data.get("dien_thoai_ben_b") or final_dynamic_data.get("dien_thoai") or ai_flat_data.get("dien_thoai", ""),
                            "nguoi_dai_dien_ben_b": final_dynamic_data.get("nguoi_dai_dien_ben_b") or ai_flat_data.get("nguoi_dai_dien_ben_b", ""),
                            "chuc_vu_ben_b": final_dynamic_data.get("chuc_vu_ben_b") or ai_flat_data.get("chuc_vu_ben_b", ""),
                            "ho_ten_nguoi_de_nghi": final_dynamic_data.get("ho_ten_nguoi_de_nghi") or final_dynamic_data.get("nguoi_de_xuat") or ai_flat_data.get("ho_ten_nguoi_de_nghi", ""),
                            "nguoi_de_xuat": final_dynamic_data.get("nguoi_de_xuat") or final_dynamic_data.get("ho_ten_nguoi_de_nghi") or ai_flat_data.get("nguoi_de_xuat", ""),
                            "bo_phan": final_dynamic_data.get("bo_phan") or final_dynamic_data.get("bo_phan_cong_tac") or ai_flat_data.get("bo_phan", ""),
                            "bo_phan_cong_tac": final_dynamic_data.get("bo_phan_cong_tac") or final_dynamic_data.get("bo_phan") or ai_flat_data.get("bo_phan", ""),
                            "chuc_vu": final_dynamic_data.get("chuc_vu") or ai_flat_data.get("chuc_vu", ""),
                            "ly_do_de_nghi": final_dynamic_data.get("ly_do_de_nghi") or final_dynamic_data.get("ly_do_thanh_toan") or ai_flat_data.get("ly_do_de_nghi", ""),
                            "ly_do_thanh_toan": final_dynamic_data.get("ly_do_thanh_toan") or final_dynamic_data.get("ly_do_de_nghi") or ai_flat_data.get("ly_do_thanh_toan", ""),
                            "hinh_thuc_thanh_toan": final_dynamic_data.get("hinh_thuc_thanh_toan") or ai_flat_data.get("hinh_thuc_thanh_toan", "Chuyển khoản"),
                            "thoi_han_hoan_ung": final_dynamic_data.get("thoi_han_hoan_ung") or ai_flat_data.get("thoi_han_hoan_ung", ""),
                            "so_hop_dong": final_dynamic_data.get("so_hop_dong") or ai_flat_data.get("so_hop_dong", ""),
                            "so_chung_tu": final_dynamic_data.get("so_chung_tu") or ai_flat_data.get("so_chung_tu", ""),
                            "so_hoa_don": final_dynamic_data.get("so_chung_tu") or ai_flat_data.get("so_chung_tu", ""),
                            "muc_dich": resolved_muc_dich,
                            "muc_dich_su_dung": resolved_muc_dich,
                            "ton_kho": resolved_ton_kho,
                            "ton_kh": resolved_ton_kho,
                            "so_don_dat_hang": final_dynamic_data.get("so_don_dat_hang") or ai_flat_data.get("so_don_dat_hang", ""),
                            "so_de_xuat": final_dynamic_data.get("so_de_xuat") or ai_flat_data.get("so_de_xuat", ""),
                            "nguoi_phu_trach": final_dynamic_data.get("nguoi_phu_trach") or ai_flat_data.get("nguoi_phu_trach", ""),
                            "email_phu_trach": final_dynamic_data.get("email_phu_trach") or ai_flat_data.get("email_phu_trach", ""),
                            "nguoi_nhan_hang": final_dynamic_data.get("nguoi_nhan_hang") or ai_flat_data.get("nguoi_nhan_hang", ""),
                            "sdt_nguoi_nhan": final_dynamic_data.get("sdt_nguoi_nhan") or ai_flat_data.get("sdt_nguoi_nhan", ""),
                            "thoi_gian_thuc_hien": final_dynamic_data.get("thoi_gian_thuc_hien") or ai_flat_data.get("thoi_gian_thuc_hien", ""),
                            "dia_diem_thuc_hien": final_dynamic_data.get("dia_diem_thuc_hien") or ai_flat_data.get("dia_diem_thuc_hien", ""),
                            "thoi_gian_bao_hanh": final_dynamic_data.get("thoi_gian_bao_hanh") or ai_flat_data.get("thoi_gian_bao_hanh", ""),
                            "ty_le_tam_ung": final_dynamic_data.get("ty_le_tam_ung") or ai_flat_data.get("ty_le_tam_ung", ""),
                            "ngay_ky": ngay_k,
                            "thang_ky": thang_k,
                            "nam_ky": nam_k,
                            "ngay_thang_nam": ngay_tn,
                            "ngay_lap_phieu": ngay_tn,
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
                            
                            safe_name_ncc = re.sub(r'[^a-zA-Z0-9_\-]', '_', resolved_ncc).strip('_')
                            zip_filename = f"Bo_Ho_So_{safe_name_ncc[:25] or 'Chung_Tu'}.zip"

                            st.session_state['ready_zip'] = {
                                "data": zip_buffer.getvalue(),
                                "filename": zip_filename,
                                "partner": resolved_ncc,
                                "total": tong_thanh_toan
                            }
                            st.toast("✅ Đã hoàn tất lập và nén trọn bộ hồ sơ!", icon="🎉")

                        except Exception as e:
                            st.error(f"Đã xảy ra lỗi trong quá trình xuất hồ sơ: {str(e)}")

            # Hiển thị nút tải về và thông tin gửi Zalo nếu đã tạo hồ sơ thành công
            if st.session_state.get('ready_zip'):
                ready_zip = st.session_state['ready_zip']
                st.markdown("<br>", unsafe_allow_html=True)
                with st.container(border=True):
                    st.success(f"✅ **ĐÃ SẴN SÀNG TẢI VỀ BỘ HỒ SƠ CHỨNG TỪ!** (Đã nén {len(selected_templates)} biểu mẫu)")
                    st.download_button(
                        label=f"💾 TẢI VỀ FILE ZIP: {ready_zip['filename']}", 
                        data=ready_zip["data"], 
                        file_name=ready_zip["filename"], 
                        mime="application/zip", 
                        type="primary",
                        use_container_width=True
                    )
                    display_zalo_message(ready_zip["partner"], ready_zip["total"])