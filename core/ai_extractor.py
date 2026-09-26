from google import genai
from PIL import Image
import json
import pandas as pd
import docx
import os
import time

def repair_truncated_json(json_str: str):
    """Cố gắng sửa chữa chuỗi JSON bị ngắt giữa chừng do kích thước output lớn (100+ mặt hàng)."""
    if not json_str:
        return None
    s = json_str.strip()
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        pass

    # Tìm vị trí đóng '}' gần nhất để chốt mảng danh_sach_hang_hoa và đóng JSON
    last_brace = s.rfind('}')
    while last_brace > 0:
        truncated = s[:last_brace + 1]
        suffixes = [
            ']}',
            ']}}',
            '], "tong_ket_tien": {}, "thong_tin_vat": {}, "thong_tin_dong": {}, "danh_sach_canh_bao": []}',
            '"]}',
            '0]}'
        ]
        for suffix in suffixes:
            try:
                candidate = truncated + suffix
                parsed = json.loads(candidate)
                if isinstance(parsed, dict) and "danh_sach_hang_hoa" in parsed:
                    return parsed
            except Exception:
                pass
        last_brace = s.rfind('}', 0, last_brace)
    return None

class AIExtractor:
    _current_key_index = 0  # Biến tĩnh lưu vị trí key hiện tại để xoay vòng liên tục

    def __init__(self, api_keys):
        """
        Khởi tạo AIExtractor với cụm API Keys xoay vòng (Round-Robin) & Chống nghẽn tải
        :param api_keys: có thể là str (1 key hoặc chuỗi key cách nhau bởi dấu phẩy) hoặc list[str]
        """
        if isinstance(api_keys, str):
            self.api_keys = [k.strip() for k in api_keys.split(",") if k.strip()]
        elif isinstance(api_keys, list):
            self.api_keys = [k.strip() for k in api_keys if isinstance(k, str) and k.strip()]
        else:
            self.api_keys = []

        # Mô hình chính thức theo khuyến nghị mới nhất của Google (tốc độ cao, chuẩn xác)
        self.primary_model = 'gemini-2.5-flash'
        # Các mô hình phao cứu sinh với hạn ngạch RPM rộng nếu model chính nghẽn tải
        self.fallback_models = ['gemini-flash-latest', 'gemini-2.5-flash-lite', 'gemini-3.1-flash-lite', 'gemini-3-flash-preview']

    def extract_invoice_data(self, file_path: str, expected_tags=None):
        if not self.api_keys:
            return {"error": "Chưa tìm thấy danh sách Gemini API Keys trong file .env hoặc cấu hình Secrets."}

        file_extension = file_path.split('.')[-1].lower()
        if file_extension not in ['png', 'jpg', 'jpeg', 'pdf', 'docx', 'xlsx']:
            return {"error": f"Định dạng file .{file_extension} chưa được hỗ trợ."}

        # Chuẩn bị trước nội dung cho ảnh / word / excel / digital pdf
        image_content = None
        text_content_payload = None

        try:
            if file_extension in ['png', 'jpg', 'jpeg']:
                with Image.open(file_path) as img:
                    image_content = img.copy()
            elif file_extension == 'pdf':
                # Thử trích xuất văn bản số hóa trực tiếp bằng PyMuPDF (cực nhanh, không tốn token hình ảnh)
                try:
                    import pymupdf
                    pdf_doc = pymupdf.open(file_path)
                    pdf_text = ""
                    for p_idx, page in enumerate(pdf_doc):
                        p_txt = page.get_text().strip()
                        if p_txt:
                            pdf_text += f"\n--- Trang {p_idx + 1} ---\n{p_txt}\n"
                    pdf_doc.close()
                    if len(pdf_text.strip()) > 50:
                        text_content_payload = f"Dữ liệu văn bản bóc tách nguyên bản từ hóa đơn điện tử (PDF):\n{pdf_text}"
                except Exception:
                    pass
            elif file_extension == 'docx':
                doc = docx.Document(file_path)
                text = "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ""])
                text_content_payload = f"Dữ liệu văn bản bóc tách từ file Word:\n{text}"
            elif file_extension == 'xlsx':
                df_dict = pd.read_excel(file_path, sheet_name=None)
                text = ""
                for sheet_name, df in df_dict.items():
                    text += f"\n--- Sheet: {sheet_name} ---\n{df.to_string(index=False)}"
                text_content_payload = f"Dữ liệu bảng tính bóc tách từ file Excel:\n{text}"
        except Exception as e:
            return {"error": f"Không thể đọc file tài liệu: {str(e)}"}

        # Chuẩn bị Prompt
        tags_instruction = f"""
        CHÚ Ý ĐẶC BIỆT CHO OBJECT 'thong_tin_dong':
        Cần tìm và trích xuất các giá trị cho danh sách thẻ sau: {list(expected_tags)}
        - QUY TẮC BẮT BUỘC: Các trường về tài khoản ngân hàng, tên đơn vị thụ hưởng, mã số thuế, địa chỉ (như 'ten_cong_ty', 'so_tai_khoan', 'ten_ngan_hang', 'ma_so_thue', 'dia_chi') PHẢI LẤY CỦA BÊN BÁN / NHÀ CUNG CẤP / ĐƠN VỊ THỤ HƯỞNG. TUYỆT ĐỐI KHÔNG LẤY THÔNG TIN CỦA BÊN MUA!
        - KỶ LUẬT THÉP: TUYỆT ĐỐI KHÔNG tự suy luận, không bịa đặt. Thông tin nào hoàn toàn không có trên tài liệu, BẮT BUỘC để chuỗi rỗng `""` để người dùng tự nhập tay.
        """ if expected_tags else ""

        prompt = f"""
        Bạn là Kế toán trưởng và Chuyên gia ERP hàng đầu. Nhiệm vụ của bạn là đọc và BÓC TÁCH CHÍNH XÁC thông tin từ tài liệu chứng từ kế toán theo đúng nguyên tắc "Chứng từ có sao ghi vậy".
        {tags_instruction}

        KỶ LUẬT THÉP VỀ TÍNH CHÍNH XÁC - TUYỆT ĐỐI KHÔNG BỊA ĐẶT (ZERO HALLUCINATION):
        1. NGUYÊN TẮC BẤT DI BẤT DỊCH: Chỉ bóc tách những gì CÓ THỰC và HIỂN THỊ RÕ RÀNG trên hóa đơn / báo giá / chứng từ.
        2. Nếu trên tài liệu KHÔNG CÓ thông tin (ví dụ: không có số tài khoản, không có tên ngân hàng, không có người đại diện, không có chức vụ, không có số hợp đồng, không có số phiếu đề xuất...):
           -> BẮT BUỘC ĐỂ CHUỖI RỖNG `""` (hoặc số 0 cho số lượng/tiền nếu không có).
           -> TUYỆT ĐỐI KHÔNG tự suy đoán, không bịa đặt, không đoán mò số tài khoản hay ngân hàng. Hệ thống có giao diện form để người dùng tự nhập tay chính xác.
        3. TUYỆT ĐỐI KHÔNG lấy số tài khoản, tên ngân hàng hoặc tên công ty của Bên Mua gán vào Bên Bán / Nhà Cung Cấp.

        QUY TẮC PHÂN BIỆT BÊN BÁN VÀ BÊN MUA (RẤT QUAN TRỌNG - KHÔNG ĐƯỢC NHẦM LẪN):
        1. BÊN BÁN (NHÀ CUNG CẤP / ĐƠN VỊ PHÁT HÀNH HÓA ĐƠN / ĐƠN VỊ THỤ HƯỞNG):
           - Là bên bán hàng hóa, cung cấp dịch vụ hoặc gửi báo giá.
           - Phải trích xuất đầy đủ: Tên công ty bên bán (`ten_cong_ty`), Mã số thuế (`ma_so_thue`), Địa chỉ (`dia_chi`), Điện thoại (`dien_thoai`), Email (`email`).
           - ĐẶC BIỆT: Trích xuất SỐ TÀI KHOẢN NGÂN HÀNG (`so_tai_khoan`) và TÊN NGÂN HÀNG (`ten_ngan_hang`) của BÊN BÁN (thường nằm ở phần đầu trang thông tin bên bán, hoặc ghi chú thông tin thanh toán chuyển khoản của bên bán để người mua chuyển tiền vào). Nếu bên bán không ghi số tài khoản trên chứng từ: BẮT BUỘC ĐỂ CHUỖI RỖNG `""`.
           - Tên người đại diện bên bán (`nguoi_dai_dien`), chức vụ (`chuc_vu`) nếu có trên tài liệu, nếu không có để chuỗi rỗng `""`.
        2. BÊN MUA (KHÁCH HÀNG):
           - Là đơn vị mua hàng (ví dụ: CÔNG TY CỔ PHẦN VIỆT NAM GRAPHENE GLOBAL, MST 4101649609, STK 884249867).
           - Ghi vào object 'thong_tin_khach_hang'.
           - TUYỆT ĐỐI KHÔNG lấy tên công ty hoặc số tài khoản của Bên Mua gán vào 'thong_tin_nha_cung_cap' hay các trường thụ hưởng.

        QUY TẮC TRÍCH XUẤT 100% DANH SÁCH HÀNG HÓA (HỖ TRỢ ĐƠN HÀNG LỚN 100+ MÃ HÀNG):
        1. Nếu tài liệu có nhiều trang (3-10 trang), bạn PHẢI quét đọc toàn bộ tất cả các trang từ trang đầu tiên đến trang cuối cùng.
        2. Trích xuất ĐẦY ĐỦ 100% từng dòng hàng hóa vào 'danh_sach_hang_hoa'.
        3. TUYỆT ĐỐI KHÔNG được tóm tắt, không dùng '...', không được bỏ qua bất kỳ dòng nào dù danh sách có 50 hay 100+ mặt hàng.
        4. Tiền tệ, số lượng là SỐ THỰC (không phẩy, không khoảng trắng, không "VNĐ").
        5. XÁC ĐỊNH THUẾ VAT: Bóc tách chính xác tỷ lệ % thuế của từng mặt hàng (0, 5, 8, 10...) vào trường `thue_suat` (số thực). Nếu không ghi riêng từng dòng thì lấy thuế suất chung của hóa đơn. Nếu hoàn toàn không có thông tin thuế, hãy để thue_suat là 0.

        CẤU TRÚC JSON BẮT BUỘC:
        {{
          "thong_tin_nha_cung_cap": {{
            "ten_cong_ty": "", 
            "dia_chi": "", 
            "dien_thoai": "", 
            "ma_so_thue": "", 
            "email": "",
            "so_tai_khoan": "",
            "ten_ngan_hang": "",
            "nguoi_dai_dien": "",
            "chuc_vu": ""
          }},
          "thong_tin_chung": {{"loai_chung_tu": "", "so_chung_tu": "", "ngay_thang_nam": ""}},
          "thong_tin_khach_hang": {{"ten_khach_hang": "", "dia_chi": "", "ma_so_thue": "", "so_tai_khoan": "", "ten_ngan_hang": ""}},
          "danh_sach_hang_hoa": [{{"stt": 1, "ten_hang_hoa": "", "don_vi_tinh": "", "so_luong": 0, "don_gia": 0, "thanh_tien": 0, "thue_suat": 0}}],
          "tong_ket_tien": {{"tong_tien_truoc_thue": 0, "thue_suat_vat": 0, "tien_thue_vat": 0, "tong_tien_thanh_toan": 0, "so_tien_viet_bang_chu": ""}},
          "thong_tin_vat": {{"da_bao_gom_vat": false, "thue_suat": 0}},
          "thong_tin_dong": {{}},
          "danh_sach_canh_bao": []
        }}
        """

        num_keys = len(self.api_keys)
        models_to_try = [self.primary_model] + self.fallback_models
        last_error_msg = ""
        hit_quota_error = False

        for model_candidate in models_to_try:
            start_index = AIExtractor._current_key_index
            
            for offset in range(num_keys):
                key_idx = (start_index + offset) % num_keys
                api_key = self.api_keys[key_idx]
                client = None
                uploaded_pdf = None

                try:
                    client = genai.Client(api_key=api_key)

                    # Chuẩn bị media_content
                    if file_extension in ['png', 'jpg', 'jpeg']:
                        media_content = image_content
                    elif text_content_payload:
                        media_content = text_content_payload
                    elif file_extension == 'pdf':
                        uploaded_pdf = client.files.upload(file=file_path)
                        media_content = uploaded_pdf
                    else:
                        media_content = text_content_payload

                    # Thực hiện gọi AI với max_output_tokens tối đa 65536 để chứa 100+ mặt hàng
                    response = client.models.generate_content(
                        model=model_candidate,
                        contents=[media_content, prompt],
                        config={
                            "response_mime_type": "application/json",
                            "max_output_tokens": 65536
                        }
                    )

                    if not response or not response.text:
                        continue

                    text_result = response.text.strip()
                    if text_result.startswith("```json"): 
                        text_result = text_result[7:-3].strip()
                    elif text_result.startswith("```"): 
                        text_result = text_result[3:-3].strip()

                    try:
                        parsed_data = json.loads(text_result)
                    except json.JSONDecodeError:
                        parsed_data = repair_truncated_json(text_result)

                    if not parsed_data or not isinstance(parsed_data, dict):
                        last_error_msg = "Không thể phân tích cấu trúc dữ liệu JSON từ phản hồi AI."
                        continue

                    # Chuẩn hóa an toàn các trường dict cấp cao để không bao giờ bị None/null
                    for key_dict in ["thong_tin_nha_cung_cap", "thong_tin_chung", "thong_tin_khach_hang", "thong_tin_vat", "thong_tin_dong", "tong_ket_tien"]:
                        if parsed_data.get(key_dict) is None or not isinstance(parsed_data.get(key_dict), dict):
                            parsed_data[key_dict] = {}
                    if parsed_data.get("danh_sach_hang_hoa") is None or not isinstance(parsed_data.get("danh_sach_hang_hoa"), list):
                        parsed_data["danh_sach_hang_hoa"] = []
                    if parsed_data.get("danh_sach_canh_bao") is None or not isinstance(parsed_data.get("danh_sach_canh_bao"), list):
                        parsed_data["danh_sach_canh_bao"] = []

                    # Thành công: Cập nhật vị trí xoay vòng cho lần gọi kế tiếp
                    AIExtractor._current_key_index = (key_idx + 1) % num_keys
                    return parsed_data

                except json.JSONDecodeError:
                    last_error_msg = "AI trả về định dạng dữ liệu không hợp lệ."
                    continue
                except Exception as e:
                    err_str = str(e)
                    err_lower = err_str.lower()
                    last_error_msg = err_str

                    # Nhận diện lỗi quá tải model 503 / 500 / Overloaded -> chuyển model ngay lập tức!
                    if "503" in err_str or "overloaded" in err_lower or "unavailable" in err_lower or "500" in err_str:
                        print(f"[AIExtractor] Model {model_candidate} đang quá tải ({err_str}). Lập tức chuyển sang mô hình phao cứu sinh tiếp theo...")
                        break

                    # Nhận diện lỗi 429 / Quota / Resource Exhausted -> thử key tiếp theo
                    if "429" in err_str or "resource_exhausted" in err_lower or "quota" in err_lower:
                        hit_quota_error = True
                        print(f"[AIExtractor] Key #{key_idx + 1}/{num_keys} đạt giới hạn quota ({model_candidate}). Đang chuyển key khác...")
                        time.sleep(1.0)  # Giãn cách 1s chống rate limit RPM
                        continue
                    else:
                        continue
                finally:
                    # Luôn xóa file PDF tạm trên máy chủ AI nếu có upload
                    if uploaded_pdf and client:
                        try:
                            client.files.delete(name=uploaded_pdf.name)
                        except Exception:
                            pass

        # Khi toàn bộ các key và các model đều hết quota
        if hit_quota_error:
            return {
                "error": "⚠️ Cụm máy chủ AI tạm thời đạt giới hạn yêu cầu (429 Quota Exceeded). "
                         "Vui lòng đợi 30 giây - 1 phút rồi thử lại."
            }
        
        return {"error": f"Không thể trích xuất tài liệu lúc này. Chi tiết: {last_error_msg}"}