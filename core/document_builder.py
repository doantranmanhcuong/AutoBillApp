from docx import Document
import openpyxl
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
        while search_key in paragraph.text:
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
                        from docx.table import _Row
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
            search_key = f"{{{key_name}}}"
            val_str = f"{value:,.0f}" if isinstance(value, (int, float)) and key_name in ["tong_tien_hang", "thue_gtgt", "tong_thanh_toan", "tong_cong"] else str(int(value) if isinstance(value, float) and value.is_integer() else (value or ""))
            
            for p in doc.paragraphs: 
                DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs: 
                            DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
            for section in doc.sections:
                for p in section.header.paragraphs:
                    DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
                for p in section.footer.paragraphs:
                    DocumentBuilder._replace_tag_in_paragraph(p, search_key, val_str)
        
        # 3. Dọn dẹp các tag chưa có dữ liệu để không xuất hiện ký tự thô
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
    def fill_excel_template(template_path, output_path, data_dict):
        wb = openpyxl.load_workbook(template_path)
        sheet = wb.active
        danh_sach = data_dict.get("danh_sach_hang_hoa", [])
        table_start_row = None
        col_map = {}
        list_keys = ["stt", "ten_hang_hoa", "muc_dich", "ton_kho", "nha_cung_cap", "don_vi_tinh", "so_luong", "don_gia", "gia_niem_yet", "gia_mua", "thanh_tien", "ghi_chu", "tong_cong"]
        
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

        # 2. Xử lý nhân bản dòng và giữ liên kết ô gộp (merged cells)
        merged_bounds = [m.bounds for m in sheet.merged_cells.ranges]
        merged_coords = [m.coord for m in list(sheet.merged_cells.ranges)]
        for coord in merged_coords:
            sheet.unmerge_cells(coord)
        
        so_dong_them = len(danh_sach) - 1 if danh_sach else 0

        if table_start_row and so_dong_them > 0:
            sheet.insert_rows(table_start_row + 1, so_dong_them)
            for i in range(1, so_dong_them + 1):
                for c in range(1, sheet.max_column + 1):
                    src_cell = sheet.cell(row=table_start_row, column=c)
                    tgt_cell = sheet.cell(row=table_start_row + i, column=c)
                    if src_cell.has_style:
                        tgt_cell.font = copy(src_cell.font)
                        tgt_cell.border = copy(src_cell.border)
                        tgt_cell.fill = copy(src_cell.fill)
                        tgt_cell.number_format = copy(src_cell.number_format)
                        tgt_cell.alignment = copy(src_cell.alignment)

        for min_col, min_row, max_col, max_row in merged_bounds:
            if table_start_row and min_row > table_start_row:
                sheet.merge_cells(start_row=min_row + so_dong_them, start_column=min_col, end_row=max_row + so_dong_them, end_column=max_col)
            elif table_start_row and min_row == table_start_row:
                for i in range(so_dong_them + 1):
                    sheet.merge_cells(start_row=min_row + i, start_column=min_col, end_row=max_row + i, end_column=max_col)
            else:
                sheet.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)

        # 3. Điền các trường thông tin đơn lẻ (quét động an toàn tối đa 250 dòng x 40 cột)
        total_scan_r = min(sheet.max_row or 1, 250)
        total_scan_c = min(sheet.max_column or 1, 40)
        for row in sheet.iter_rows(max_row=total_scan_r, max_col=total_scan_c):
            for cell in row:
                if isinstance(cell.value, str):
                    for key, value in data_dict.items():
                        if isinstance(value, list):
                            continue
                        if f"{{{key}}}" in cell.value:
                            val_str = f"{value:,.0f}" if isinstance(value, (int, float)) and key in ["tong_tien_hang", "thue_gtgt", "tong_thanh_toan", "tong_cong"] else str(int(value) if isinstance(value, float) and value.is_integer() else (value or ""))
                            cell.value = cell.value.replace(f"{{{key}}}", val_str)

        # 4. Điền dữ liệu vào bảng kê hàng hóa
        if table_start_row and col_map:
            for idx, item in enumerate(danh_sach):
                for key_name, col_idx in col_map.items():
                    val = idx + 1 if key_name == "stt" else (float(item.get(key_name, 0) or 0) if key_name in ["so_luong", "don_gia", "thanh_tien", "tong_cong"] else item.get(key_name, ""))
                    c = sheet.cell(row=table_start_row + idx, column=col_idx)
                    c.value = val
                    if key_name in ["so_luong", "don_gia", "thanh_tien", "tong_cong"]:
                        c.number_format = '#,##0'

        # 5. Dọn dẹp các tag chưa có dữ liệu trong Excel
        for row in sheet.iter_rows(max_row=total_scan_r, max_col=total_scan_c):
            for cell in row:
                if isinstance(cell.value, str) and "{" in cell.value:
                    cleaned = re.sub(r'\{[a-zA-Z0-9_]+\}', '', cell.value).strip()
                    cell.value = cleaned
                    
        wb.save(output_path)
        return True