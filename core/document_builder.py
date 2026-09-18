from docx import Document
from docx.table import _Row
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

        # 1. Nhân bản và điền bảng kê hàng hóa
        if danh_sach:
            for table in doc.tables:
                template_row = next((r for r in table.rows if "{ten_hang_hoa}" in "".join(c.text for c in r.cells) or "{stt}" in "".join(c.text for c in r.cells)), None)
                if template_row:
                    parent_tbl = template_row._tr.getparent()
                    for idx, item in enumerate(danh_sach):
                        new_tr = deepcopy(template_row._tr)
                        parent_tbl.insert(parent_tbl.index(template_row._tr), new_tr)
                        new_row = _Row(new_tr, table)
                        
                        for cell in new_row.cells:
                            for p in cell.paragraphs:
                                if "{stt}" in p.text: 
                                    DocumentBuilder._replace_tag_in_paragraph(p, "{stt}", str(idx + 1))
                                for k, v in item.items():
                                    search_key = f"{{{k}}}"
                                    if search_key in p.text:
                                        val_str = f"{v:,.0f}" if isinstance(v, (int, float)) and k in ["so_luong", "don_gia", "thanh_tien", "tong_cong"] else str(v if v is not None else "")
                                        DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
                                
                                for run in p.runs:
                                    run.text = re.sub(r'\{[a-zA-Z0-9_]+\}', '', run.text)
                                    
                    parent_tbl.remove(template_row._tr)

        # 2. Điền các trường thông tin đơn lẻ (header, footer, thông tin chung)
        for key_name, value in data_dict.items():
            if isinstance(value, list):
                continue
            money_keys = [
                "tong_tien_hang", "thue_gtgt", "tong_thanh_toan", "tong_cong",
                "thue_5", "thue_gtgt_5", "thue_8", "thue_gtgt_8", "thue_10", "thue_gtgt_10",
                "tien_hang_5", "tien_hang_8", "tien_hang_10"
            ]
            val_str = f"{value:,.0f}" if isinstance(value, (int, float)) and key_name in money_keys else str(int(value) if isinstance(value, float) and value.is_integer() else (value or ""))
            
            search_key = f"{{{key_name}}}"
            for p in doc.paragraphs: 
                if search_key in p.text:
                    DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs: 
                            if search_key in p.text:
                                DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
            for section in doc.sections:
                for p in section.header.paragraphs:
                    if search_key in p.text:
                        DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
                for p in section.footer.paragraphs:
                    if search_key in p.text:
                        DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
        # 3. Tự động cập nhật nhãn Thuế GTGT nếu có phân rã nhiều mức thuế
        breakdown_str = data_dict.get("chi_tiet_thue")
        rate = data_dict.get("thue_suat")
        if breakdown_str or rate is not None:
            def update_vat_label(p):
                if "Thuế GTGT" in p.text and ("%" in p.text or ":" in p.text):
                    if breakdown_str:
                        for target in ["Thuế GTGT 8%:", "Thuế GTGT 8%", "Thuế GTGT:"]:
                            if target in p.text:
                                DocumentBuilder._replace_tag_in_paragraph(p, target, f"Thuế GTGT ({breakdown_str}):")
                    elif rate is not None:
                        try:
                            r_num = float(rate)
                            for target in ["Thuế GTGT 8%:", "Thuế GTGT 8%"]:
                                if target in p.text:
                                    DocumentBuilder._replace_tag_in_paragraph(p, target, f"Thuế GTGT {r_num:g}%:")
                        except Exception:
                            pass

            for p in doc.paragraphs:
                update_vat_label(p)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            update_vat_label(p)

        # 4. Dọn dẹp các tag chưa có dữ liệu để không xuất hiện ký tự thô
        tag_pattern = re.compile(r'\{([a-zA-Z0-9_]+)\}')
        def clean_tags_in_p(p):
            for tag in tag_pattern.findall(p.text):
                DocumentBuilder._replace_tag_in_paragraph(p, f"{{{tag}}}", "")

        for p in doc.paragraphs:
            clean_tags_in_p(p)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        clean_tags_in_p(p)
        for section in doc.sections:
            for p in section.header.paragraphs:
                clean_tags_in_p(p)
            for p in section.footer.paragraphs:
                clean_tags_in_p(p)
        
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
            for i in range(1, so_dong_them + 1):
                for c in range(1, sheet.max_column + 1):
                    src_cell = sheet.cell(row=table_start_row, column=c)
                    tgt_cell = sheet.cell(row=table_start_row + i, column=c)
                    if src_cell.has_style:
                        tgt_cell.font = copy(src_cell.font)
                        tgt_cell.fill = copy(src_cell.fill)
                        tgt_cell.number_format = copy(src_cell.number_format)
                        tgt_cell.alignment = copy(src_cell.alignment)

            # Khôi phục các ô gộp với chỉ số hàng đã được dịch chuyển chuẩn xác
            for min_col, min_row, max_col, max_row in merged_bounds:
                if min_row > table_start_row:
                    sheet.merge_cells(start_row=min_row + so_dong_them, start_column=min_col, end_row=max_row + so_dong_them, end_column=max_col)
                elif min_row == table_start_row:
                    for i in range(so_dong_them + 1):
                        sheet.merge_cells(start_row=min_row + i, start_column=min_col, end_row=max_row + i, end_column=max_col)
                else:
                    sheet.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

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

        # 3. Điền các trường thông tin đơn lẻ (quét động an toàn tối đa 250 dòng x 40 cột)
        total_scan_r = min(sheet.max_row or 1, 250)
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
                            if val_num.is_integer():
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