from docx import Document
from docx.table import _Row
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import openpyxl
from openpyxl.styles import Border, Side, Alignment, Font
from copy import copy, deepcopy
import re

class DocumentBuilder:
    @staticmethod
    def _replace_tag_in_paragraph(paragraph, search_key, val_str):
        """
        Thuật toán Run-Span Replacement:
        Bảo toàn 100% định dạng (Font, Cỡ chữ, Màu sắc, In đậm/nghiêng) của template
        ngay cả khi Word tự động ngắt tag {tag} qua nhiều run khác nhau.
        """
        if search_key not in paragraph.text:
            return

        # Trường hợp 1: Tag nằm trọn vẹn trong 1 run đơn lẻ
        for run in paragraph.runs:
            if search_key in run.text:
                run.text = run.text.replace(search_key, val_str)
                if search_key not in paragraph.text:
                    return

        # Trường hợp 2: Tag bị chia cắt qua nhiều run liên tiếp
        loop_guard = 0
        while search_key in paragraph.text and loop_guard < 50:
            loop_guard += 1
            full_text = "".join(r.text for r in paragraph.runs)
            if search_key not in full_text:
                break
            start_pos = full_text.find(search_key)
            end_pos = start_pos + len(search_key)

            curr = 0
            start_run_idx = None
            end_run_idx = None
            start_offset = 0
            end_offset = 0

            for i, r in enumerate(paragraph.runs):
                run_len = len(r.text)
                if start_run_idx is None and curr + run_len > start_pos:
                    start_run_idx = i
                    start_offset = start_pos - curr
                if end_run_idx is None and curr + run_len >= end_pos:
                    end_run_idx = i
                    end_offset = end_pos - curr
                    break
                curr += run_len

            if start_run_idx is None or end_run_idx is None:
                break

            if start_run_idx == end_run_idx:
                r = paragraph.runs[start_run_idx]
                r.text = r.text[:start_offset] + val_str + r.text[end_offset:]
            else:
                prefix = paragraph.runs[start_run_idx].text[:start_offset]
                suffix = paragraph.runs[end_run_idx].text[end_offset:]
                # Giữ nguyên kiểu định dạng của start_run (chứa tag)
                paragraph.runs[start_run_idx].text = prefix + val_str
                # Xóa trắng các run trung gian mà không làm biến đổi cấu trúc
                for m in range(start_run_idx + 1, end_run_idx):
                    paragraph.runs[m].text = ""
                paragraph.runs[end_run_idx].text = suffix

    @staticmethod
    def fill_word_template(template_path, output_path, data_dict):
        doc = Document(template_path)
        danh_sach = data_dict.get("danh_sach_hang_hoa", [])
        breakdown_str = data_dict.get("chi_tiet_thue")
        rate = data_dict.get("thue_suat")
        tien_thue_tong = data_dict.get("thue_gtgt", 0)
        try:
            tien_thue_num = float(tien_thue_tong or 0)
        except Exception:
            tien_thue_num = 0.0

        # Chuẩn bị nhãn thuế GTGT mặc định
        if breakdown_str:
            vat_label_text = f"Tiền thuế GTGT ({breakdown_str})"
            contract_vat_label = f"Thuế GTGT ({breakdown_str})"
        elif rate is not None:
            try:
                r_val = float(rate)
                vat_label_text = f"Tiền thuế GTGT ({r_val:g}%):"
                contract_vat_label = f"Thuế VAT {r_val:g}%"
            except Exception:
                vat_label_text = f"Tiền thuế GTGT ({rate}):"
                contract_vat_label = f"Thuế VAT {rate}"
        else:
            vat_label_text = "Tiền thuế GTGT:"
            contract_vat_label = "Thuế VAT"

        if "thue_suat_label" not in data_dict:
            data_dict["thue_suat_label"] = contract_vat_label

        # 1. Nhân bản và điền bảng kê hàng hóa
        if danh_sach:
            for table in doc.tables:
                # Chuẩn hóa các header bị cứng số 8% (ví dụ: 'Thuế suất 8%')
                for row in table.rows[:2]:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if "Thuế suất 8%" in p.text:
                                DocumentBuilder._replace_tag_in_paragraph(p, "Thuế suất 8%", "Thuế suất")

                template_row = next((r for r in table.rows if "{ten_hang_hoa}" in "".join(c.text for c in r.cells) or "{stt}" in "".join(c.text for c in r.cells)), None)
                if template_row:
                    parent_tbl = template_row._tr.getparent()
                    tpl_row_idx = parent_tbl.index(template_row._tr)

                    # Kiểm tra xem bảng này đã có ô hoặc dòng riêng cho thuế ({thue_gtgt}) chưa
                    table_full_text = "".join(c.text for r in table.rows for c in r.cells)
                    has_vat_row_already = "{thue_gtgt}" in table_full_text or "Thuế VAT" in table_full_text or "Thuế GTGT" in table_full_text

                    for idx, item in enumerate(danh_sach):
                        new_tr = deepcopy(template_row._tr)
                        parent_tbl.insert(tpl_row_idx + idx, new_tr)
                        new_row = _Row(new_tr, table)
                        
                        item_rate = item.get("thue_suat", rate if rate is not None else 8)
                        rate_display = f"{item_rate:g}%" if isinstance(item_rate, (int, float)) else str(item_rate)
                        if "%" not in rate_display:
                            rate_display += "%"

                        for cell in new_row.cells:
                            for p in cell.paragraphs:
                                if "{stt}" in p.text: 
                                    DocumentBuilder._replace_tag_in_paragraph(p, "{stt}", str(idx + 1))
                                # Tự động thay thế ô cứng '8%' trong bảng thành thuế suất thực tế của món hàng
                                if p.text.strip() == "8%" and "{thue_suat}" not in p.text:
                                    DocumentBuilder._replace_tag_in_paragraph(p, "8%", rate_display)
                                if "{thue_suat}" in p.text:
                                    DocumentBuilder._replace_tag_in_paragraph(p, "{thue_suat}", rate_display)

                                if "{" in p.text:
                                    for k, v in item.items():
                                        search_key = f"{{{k}}}"
                                        if search_key in p.text:
                                            if k == "so_luong" and isinstance(v, (int, float)):
                                                is_int = isinstance(v, int) or (isinstance(v, float) and v.is_integer())
                                                val_str = f"{int(v):,}" if is_int else f"{v:g}"
                                            elif isinstance(v, (int, float)) and k in ["don_gia", "thanh_tien", "tong_cong", "gia_niem_yet", "gia_mua", "tien_thue"]:
                                                val_str = f"{v:,.0f}"
                                            else:
                                                val_str = str(v if v is not None else "")
                                            DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
                                    
                                    for run in p.runs:
                                        run.text = re.sub(r'\{[a-zA-Z0-9_]+\}', '', run.text)

                    # BỔ SUNG DÒNG THUẾ GTGT CHO BẢNG KÊ THANH TOÁN (Nếu bảng chưa có dòng thuế và có phát sinh tiền thuế)
                    if not has_vat_row_already and tien_thue_num > 0 and len(template_row.cells) >= 3:
                        vat_tr = deepcopy(template_row._tr)
                        parent_tbl.insert(tpl_row_idx + len(danh_sach), vat_tr)
                        vat_row = _Row(vat_tr, table)

                        # Tìm đúng vị trí cột thành tiền của dòng mẫu để điền số tiền thuế
                        thanh_tien_col_idx = next((c_i for c_i, c in enumerate(template_row.cells) if "{thanh_tien}" in c.text), len(template_row.cells) - 2)

                        for c_idx, cell in enumerate(vat_row.cells):
                            for p in cell.paragraphs:
                                # Xóa sạch các run cũ trong dòng mẫu clone
                                for r in p.runs:
                                    r.text = ""
                                if c_idx == 0:
                                    p.text = ""
                                elif c_idx == 1:
                                    p.text = vat_label_text
                                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                                    if p.runs:
                                        p.runs[0].font.name = "Times New Roman"
                                        p.runs[0].font.size = Pt(11)
                                        p.runs[0].font.italic = True
                                elif c_idx == thanh_tien_col_idx:
                                    p.text = f"{tien_thue_num:,.0f}"
                                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                                    if p.runs:
                                        p.runs[0].font.name = "Times New Roman"
                                        p.runs[0].font.size = Pt(11)
                                        p.runs[0].font.bold = True
                                else:
                                    p.text = ""
                                    
                    parent_tbl.remove(template_row._tr)

        # 2. Điền các trường thông tin đơn lẻ và chuẩn hóa câu chữ (Tối ưu hóa 1-Pass siêu tốc)
        replacements = {}
        money_keys = {
            "tong_tien_hang", "thue_gtgt", "tong_thanh_toan", "tong_cong",
            "thue_5", "thue_gtgt_5", "thue_8", "thue_gtgt_8", "thue_10", "thue_gtgt_10",
            "tien_hang_5", "tien_hang_8", "tien_hang_10"
        }
        for key_name, value in data_dict.items():
            if isinstance(value, list):
                continue
            val_str = f"{value:,.0f}" if isinstance(value, (int, float)) and key_name in money_keys else str(int(value) if isinstance(value, float) and value.is_integer() else (value or ""))
            replacements[f"{{{key_name}}}"] = val_str

        tag_pattern = re.compile(r'\{([a-zA-Z0-9_]+)\}')

        def process_doc_paragraph(p):
            p_text = p.text
            # 1. Thay thế tag nếu có dấu {
            if "{" in p_text:
                for search_key, val_str in replacements.items():
                    if search_key in p.text:
                        DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
                # Dọn dẹp các tag chưa có dữ liệu để không xuất hiện ký tự thô
                if "{" in p.text:
                    for tag in tag_pattern.findall(p.text):
                        DocumentBuilder._replace_tag_in_paragraph(p, f"{{{tag}}}", "")

            # 2. Thay thế linh hoạt câu điều khoản thuế trong hợp đồng: 'thuế giá trị gia tăng 8% (tám phần trăm)'
            if "thuế giá trị gia tăng 8% (tám phần trăm)" in p.text:
                if breakdown_str:
                    target_clause = f"thuế giá trị gia tăng ({breakdown_str})"
                elif rate is not None:
                    try:
                        r_val = float(rate)
                        r_words = {0: "không", 5: "năm", 8: "tám", 10: "mười"}.get(int(r_val), str(r_val))
                        target_clause = f"thuế giá trị gia tăng {r_val:g}% ({r_words} phần trăm)"
                    except Exception:
                        target_clause = f"thuế giá trị gia tăng {rate}"
                else:
                    target_clause = "thuế giá trị gia tăng"
                DocumentBuilder._replace_tag_in_paragraph(p, "thuế giá trị gia tăng 8% (tám phần trăm)", target_clause)

            # 3. Tự động cập nhật nhãn Thuế GTGT / Thuế VAT trên bảng hợp đồng hoặc văn bản
            if any(k in p.text for k in ["Thuế VAT", "Thuế GTGT", "Thuế suất", "VAT 8%"]):
                targets = [
                    "Thuế VAT 8%:", "Thuế VAT 8%", "Thuế VAT:",
                    "Thuế GTGT 8%:", "Thuế GTGT 8%", "Thuế GTGT:",
                    "Thuế suất 8%:", "Thuế suất 8%"
                ]
                for target in targets:
                    if target in p.text:
                        rep_val = f"{contract_vat_label}:" if target.endswith(":") else contract_vat_label
                        DocumentBuilder._replace_tag_in_paragraph(p, target, rep_val)
                        break

        for p in doc.paragraphs:
            process_doc_paragraph(p)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        process_doc_paragraph(p)
        for section in doc.sections:
            for p in section.header.paragraphs:
                process_doc_paragraph(p)
            for p in section.footer.paragraphs:
                process_doc_paragraph(p)
        
        doc.save(output_path)
        return True

    @staticmethod
    def _parse_number(val):
        from core.utils.helpers import safe_float
        return safe_float(val, 0.0)

    @staticmethod
    def fill_excel_template(template_path, output_path, data_dict):
        wb = openpyxl.load_workbook(template_path)
        sheet = wb.active
        # Tự động quét tìm sheet chứa biểu mẫu (chứa ký tự '{') nếu sheet active không chứa tag
        has_tags = any('{' in str(sheet.cell(row=r, column=c).value or '') 
                       for r in range(1, min(sheet.max_row or 1, 50)) 
                       for c in range(1, min(sheet.max_column or 1, 25)))
        if not has_tags:
            for s in wb.worksheets:
                if any('{' in str(s.cell(row=r, column=c).value or '') 
                       for r in range(1, min(s.max_row or 1, 50)) 
                       for c in range(1, min(s.max_column or 1, 25))):
                    sheet = s
                    break

        danh_sach = data_dict.get("danh_sach_hang_hoa", [])
        table_start_row = None
        col_map = {}
        list_keys = [
            "stt", "ten_hang_hoa", "muc_dich", "muc_dich_su_dung", "ton_kho", "ton_kh",
            "nha_cung_cap", "don_vi_tinh", "so_luong", "don_gia", "gia_niem_yet", "gia_mua",
            "thanh_tien", "ghi_chu", "tong_cong", "thue_suat", "tien_thue"
        ]
        
        # 1. Tìm vị trí hàng mẫu bảng kê (quét thông minh trong phạm vi 80 dòng đầu và 35 cột)
        max_scan_r = min(sheet.max_row or 1, 80)
        max_scan_c = min(sheet.max_column or 1, 35)
        for r in range(1, max_scan_r + 1):
            for c in range(1, max_scan_c + 1):
                try:
                    val = sheet.cell(row=r, column=c).value
                    if val and isinstance(val, str):
                        for key in list_keys:
                            if f"{{{key}}}" in val:
                                col_map[key] = c
                                table_start_row = r
                except Exception:
                    pass
            if table_start_row:
                break

        # Bổ sung map cột từ dòng tiêu đề (nếu trong dòng mẫu chưa có thẻ tag, ví dụ 'Tồn kho', 'Đơn giá Niêm Yết')
        if table_start_row and table_start_row > 1:
            header_r = table_start_row - 1
            for c in range(1, max_scan_c + 1):
                h_val = str(sheet.cell(row=header_r, column=c).value or "").strip().lower()
                if not h_val:
                    continue
                if "tồn kho" in h_val and "ton_kho" not in col_map:
                    col_map["ton_kho"] = c
                elif "niêm yết" in h_val and "gia_niem_yet" not in col_map:
                    col_map["gia_niem_yet"] = c
                elif "mục đích" in h_val and "muc_dich" not in col_map:
                    col_map["muc_dich"] = c
                elif "nhà cung cấp" in h_val and "nha_cung_cap" not in col_map:
                    col_map["nha_cung_cap"] = c
                elif ("thuế" in h_val or "vat" in h_val) and "thue_suat" not in col_map:
                    col_map["thue_suat"] = c

        # 2. Xử lý nhân bản dòng, giữ liên kết ô gộp và đồng bộ chuẩn tỷ lệ hàng (row dimensions)
        so_dong_them = len(danh_sach) - 1 if danh_sach else 0

        if table_start_row and so_dong_them > 0:
            # Lưu lại thông số kích thước các hàng ban đầu để dịch chuyển chính xác cho phần footer
            orig_row_dimensions = {}
            for r, rd in sheet.row_dimensions.items():
                orig_row_dimensions[r] = {
                    "height": rd.height,
                    "hidden": rd.hidden
                }

            merged_bounds = [m.bounds for m in sheet.merged_cells.ranges]
            merged_coords = [m.coord for m in list(sheet.merged_cells.ranges)]
            for coord in merged_coords:
                sheet.unmerge_cells(coord)

            # Chèn các dòng mới cho danh sách hàng hóa
            sheet.insert_rows(table_start_row + 1, so_dong_them)

            # Sao chép kiểu dáng ô (font, fill, alignment, number_format) từ hàng mẫu
            # Giới hạn tối đa 40 cột an toàn, chống nghẽn khi gặp file Excel bị format cả dòng 16.384 cột
            max_col_safe = min(sheet.max_column or 1, 40)
            for i in range(1, so_dong_them + 1):
                for c in range(1, max_col_safe + 1):
                    src_cell = sheet.cell(row=table_start_row, column=c)
                    tgt_cell = sheet.cell(row=table_start_row + i, column=c)
                    if src_cell.has_style:
                        tgt_cell.font = copy(src_cell.font)
                        tgt_cell.fill = copy(src_cell.fill)
                        tgt_cell.number_format = copy(src_cell.number_format)
                        tgt_cell.alignment = copy(src_cell.alignment)
                        tgt_cell.border = copy(src_cell.border)

            # Khôi phục các ô gộp với chỉ số hàng đã được dịch chuyển chuẩn xác
            for min_col, min_row, max_col, max_row in merged_bounds:
                try:
                    if min_row > table_start_row:
                        sheet.merge_cells(start_row=min_row + so_dong_them, start_column=min_col, end_row=max_row + so_dong_them, end_column=max_col)
                    elif min_row == table_start_row:
                        for i in range(so_dong_them + 1):
                            sheet.merge_cells(start_row=min_row + i, start_column=min_col, end_row=max_row + i, end_column=max_col)
                    else:
                        sheet.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
                except Exception:
                    pass

            # Tái tạo và dịch chuyển kích thước hàng (row heights) để không bị xô lệch form
            for r in list(sheet.row_dimensions.keys()):
                if r >= table_start_row:
                    del sheet.row_dimensions[r]

            # Thiết lập chiều cao chuẩn cân đối cho từng dòng hàng hóa
            for idx, item in enumerate(danh_sach):
                curr_r = table_start_row + idx
                name_len = len(str(item.get("ten_hang_hoa") or ""))
                # Tự động tính chiều cao hàng: thoáng đãng, cân xứng, không đè viền
                if name_len > 70:
                    h = 36.0
                elif name_len > 35:
                    h = 28.0
                else:
                    h = 23.0
                sheet.row_dimensions[curr_r].height = h

            # Dịch chuyển kích thước cho tất cả các hàng bên dưới bảng (Cộng, Thuế, Chữ ký...)
            for orig_r, meta in orig_row_dimensions.items():
                if orig_r > table_start_row:
                    new_r = orig_r + so_dong_them
                    if meta["height"] is not None:
                        sheet.row_dimensions[new_r].height = meta["height"]
                    sheet.row_dimensions[new_r].hidden = meta["hidden"]
        elif table_start_row and len(danh_sach) == 1:
            # Nếu chỉ có 1 mặt hàng, chuẩn hóa chiều cao hàng mẫu
            name_len = len(str(danh_sach[0].get("ten_hang_hoa") or ""))
            sheet.row_dimensions[table_start_row].height = 28.0 if name_len > 35 else 23.0

        # 3. Điền các trường thông tin đơn lẻ (quét động an toàn tối đa 2000 dòng x 40 cột)
        total_scan_r = min(sheet.max_row or 1, 2000)
        total_scan_c = min(sheet.max_column or 1, 40)
        money_keys = [
            "tong_tien_hang", "thue_gtgt", "tong_thanh_toan", "tong_cong",
            "thue_5", "thue_gtgt_5", "thue_8", "thue_gtgt_8", "thue_10", "thue_gtgt_10",
            "tien_hang_5", "tien_hang_8", "tien_hang_10"
        ]

        for row in sheet.iter_rows(max_row=total_scan_r, max_col=total_scan_c):
            for cell in row:
                if isinstance(cell.value, str):
                    # Sửa lỗi dính chữ trong mẫu (ví dụ: 'đơn đặt hàng{so_don_dat_hang}')
                    if "hàng{so_don_dat_hang}" in cell.value:
                        cell.value = cell.value.replace("hàng{so_don_dat_hang}", "hàng {so_don_dat_hang}")

                    # Điền bổ sung các ô trống không có thẻ tag trong ĐĐH Mẫu.xlsx
                    s_val = cell.value.strip()
                    if s_val == "Ngày:" or s_val == "Ngày :":
                        ngay_val = data_dict.get("ngay_thang_nam") or data_dict.get("ngay_lap_phieu") or ""
                        if ngay_val:
                            cell.value = f"Ngày: {ngay_val}"
                    elif (s_val.startswith("Tới") and s_val.endswith(":")) or s_val == "Tới:":
                        ncc_val = data_dict.get("ten_cong_ty") or data_dict.get("nha_cung_cap") or ""
                        if ncc_val:
                            cell.value = f"Tới   : {ncc_val}"
                    elif (s_val.startswith("V/v") and s_val.endswith(":")) or s_val == "V/v:":
                        sddh = data_dict.get("so_don_dat_hang", "")
                        so_hd = data_dict.get("so_chung_tu", "")
                        ref = sddh or so_hd
                        if ref:
                            cell.value = f"V/v   : Đơn đặt hàng {ref}"

                    # Tự động cập nhật nhãn thuế nếu hóa đơn có nhiều mức thuế hoặc thuế suất cụ thể khác 8%
                    if "Thuế GTGT" in cell.value and ("%" in cell.value or ":" in cell.value):
                        breakdown_str = data_dict.get("chi_tiet_thue")
                        rate = data_dict.get("thue_suat")
                        if breakdown_str:
                            cell.value = f"Thuế GTGT ({breakdown_str}):"
                        elif rate is not None:
                            try:
                                r_num = float(rate)
                                cell.value = re.sub(r'Thuế GTGT(\s*\d+%)?', f'Thuế GTGT {r_num:g}%', cell.value)
                            except Exception:
                                pass

                    for key, value in data_dict.items():
                        if not isinstance(cell.value, str):
                            break
                        if isinstance(value, list):
                            continue
                        tag = f"{{{key}}}"
                        if tag in cell.value:
                            # Nếu toàn bộ ô chỉ chứa đúng tag số tiền -> Lưu dạng SỐ THỰC kèm number_format
                            # để người dùng tính toán được và không bị tam giác xanh (Number Stored as Text)
                            if cell.value.strip() == tag and key in money_keys:
                                cell.value = DocumentBuilder._parse_number(value)
                                cell.number_format = '#,##0'
                                cell.alignment = Alignment(horizontal='right', vertical='center')
                                break
                            else:
                                val_str = f"{value:,.0f}" if isinstance(value, (int, float)) and key in money_keys else str(int(value) if isinstance(value, float) and value.is_integer() else (value or ""))
                                cell.value = cell.value.replace(tag, val_str)
                                if "Ngày: ngày " in cell.value:
                                    cell.value = cell.value.replace("Ngày: ngày ", "Ngày: ")
                                elif "Ngày : ngày " in cell.value:
                                    cell.value = cell.value.replace("Ngày : ngày ", "Ngày: ")

        # 4. Điền dữ liệu vào bảng kê hàng hóa & Kẻ viền sắc nét chuẩn mực
        if table_start_row and col_map:
            thin_side = Side(style='thin', color='000000')
            table_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            
            min_col_table = min(col_map.values())
            max_col_table = max(col_map.values())

            # Chuẩn hóa viền dòng tiêu đề bảng kê (header row)
            if table_start_row > 1:
                h_row = table_start_row - 1
                for c_idx in range(min_col_table, max_col_table + 1):
                    h_cell = sheet.cell(row=h_row, column=c_idx)
                    h_cell.border = table_border

            for idx, item in enumerate(danh_sach):
                curr_row = table_start_row + idx

                # Kẻ viền 4 cạnh sắc nét cho tất cả các ô trong dòng bảng kê
                for c_idx in range(min_col_table, max_col_table + 1):
                    cell = sheet.cell(row=curr_row, column=c_idx)
                    cell.border = table_border

                for key_name, col_idx in col_map.items():
                    c = sheet.cell(row=curr_row, column=col_idx)
                    
                    if key_name == "stt":
                        c.value = idx + 1
                        c.alignment = Alignment(horizontal='center', vertical='center')
                    elif key_name == "thue_suat":
                        c.value = str(item.get("thue_suat") or "")
                        c.alignment = Alignment(horizontal='center', vertical='center')
                    elif key_name in ["so_luong", "don_gia", "thanh_tien", "tong_cong", "gia_niem_yet", "gia_mua", "tien_thue"]:
                        val_raw = item.get(key_name)
                        if key_name == "gia_niem_yet" and (val_raw is None or val_raw == "" or val_raw == 0 or val_raw == "0"):
                            val_raw = item.get("don_gia", 0)
                        val_num = DocumentBuilder._parse_number(val_raw)
                        
                        if key_name == "so_luong":
                            is_int = isinstance(val_num, int) or (isinstance(val_num, float) and val_num.is_integer())
                            if is_int:
                                c.value = int(val_num)
                                c.number_format = '#,##0'
                            else:
                                c.value = val_num
                                c.number_format = '#,##0.##'
                            c.alignment = Alignment(horizontal='center', vertical='center')
                        elif key_name in ["thanh_tien", "tong_cong"]:
                            c.value = val_num
                            c.number_format = '#,##0'
                            c.alignment = Alignment(horizontal='right', vertical='center')
                            f_name = c.font.name if c.font and c.font.name else 'Times New Roman'
                            f_size = c.font.size if c.font and c.font.size else 11
                            c.font = Font(name=f_name, size=f_size, bold=True)
                        else:
                            c.value = val_num
                            c.number_format = '#,##0'
                            c.alignment = Alignment(horizontal='right', vertical='center')
                    else:
                        val_str = item.get(key_name, "")
                        # Tự động gán fallback từ thông tin chung nếu từng dòng hàng hóa chưa có Nhà cung cấp hoặc Mục đích
                        if key_name == "nha_cung_cap" and not val_str:
                            val_str = data_dict.get("nha_cung_cap") or data_dict.get("ten_cong_ty") or data_dict.get("ten_cong_ty_ben_b") or ""
                        elif key_name in ["muc_dich", "muc_dich_su_dung"] and not val_str:
                            val_str = data_dict.get("muc_dich") or data_dict.get("muc_dich_su_dung") or data_dict.get("ly_do_de_nghi") or ""
                        elif key_name in ["ton_kho", "ton_kh"] and not val_str:
                            val_str = data_dict.get("ton_kho") or data_dict.get("ton_kh") or "0"

                        c.value = str(val_str or "")
                        if key_name in ["don_vi_tinh", "ton_kho", "ton_kh"]:
                            c.alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

        # 5. Dọn dẹp các tag chưa có dữ liệu trong Excel
        for row in sheet.iter_rows(max_row=total_scan_r, max_col=total_scan_c):
            for cell in row:
                if isinstance(cell.value, str) and "{" in cell.value:
                    cleaned = re.sub(r'\{[a-zA-Z0-9_]+\}', '', cell.value).strip()
                    cell.value = cleaned
                    
        wb.save(output_path)
        wb.close()
        return True