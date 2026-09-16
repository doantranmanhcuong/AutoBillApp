import re
from docx import Document
import openpyxl

def doc_so_tien_vn(n):
    if n == 0: return "Không đồng"
    units = ["", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ"]
    words = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]

    def doc_3_so(num, read_zero_hundred=False):
        h = num // 100
        t = (num % 100) // 10
        u = num % 10
        res = ""
        if h > 0 or read_zero_hundred: res += words[h] + " trăm "
        if t > 1:
            res += words[t] + " mươi "
            if u == 1: res += "mốt "
            elif u == 5: res += "lăm "
            elif u > 0: res += words[u] + " "
        elif t == 1:
            res += "mười "
            if u == 5: res += "lăm "
            elif u > 0: res += words[u] + " "
        elif t == 0 and u > 0 and (h > 0 or read_zero_hundred): res += "lẻ " + words[u] + " "
        elif t == 0 and u > 0: res += words[u] + " "
        return res.strip()

    s = ""
    group = 0
    try:
        n_int = int(round(float(n)))
    except (ValueError, TypeError):
        return "Không đồng"
        
    if n_int == 0:
        return "Không đồng"
        
    while n_int > 0:
        chunk = n_int % 1000
        n_int = n_int // 1000
        if chunk > 0:
            chunk_str = doc_3_so(chunk, read_zero_hundred=(n_int > 0))
            s = chunk_str + units[group] + " " + s
        group += 1
    
    return s.strip().capitalize() + " đồng chẵn."

def exhaustive_extract_tags(file_path):
    """Quét toàn diện các thẻ {tag}, {{tag}}, {{{ tag }}} trong tệp docx và xlsx"""
    tags = set()
    pattern = re.compile(r'\{+\s*([a-zA-Z0-9_]+)\s*\}+')
    try:
        if file_path.endswith('.docx'):
            doc = Document(file_path)
            for p in doc.paragraphs:
                tags.update(pattern.findall(p.text))
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            tags.update(pattern.findall(p.text))
        elif file_path.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file_path, data_only=True)
            for sheet in wb.worksheets:
                max_r = min(sheet.max_row or 1, 250)
                max_c = min(sheet.max_column or 1, 40)
                for row in sheet.iter_rows(max_row=max_r, max_col=max_c, values_only=True):
                    for cell in row:
                        if isinstance(cell, str):
                            tags.update(pattern.findall(cell))
    except Exception:
        pass
        
    return {t.strip().lower() for t in tags}