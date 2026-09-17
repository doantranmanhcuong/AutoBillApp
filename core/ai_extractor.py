from google import genai
from PIL import Image
import json
import pandas as pd
import docx
import os

class AIExtractor:
    _current_key_index = 0  # Biến tĩnh lưu vị trí key hiện tại để xoay vòng liên tục

    def __init__(self, api_keys):
        """
        Khởi tạo AIExtractor với cụm 3 API Keys xoay vòng (Round-Robin)
        :param api_keys: có thể là str (1 key hoặc chuỗi key cách nhau bởi dấu phẩy) hoặc list[str]
        """
        if isinstance(api_keys, str):
            self.api_keys = [k.strip() for k in api_keys.split(",") if k.strip()]
        elif isinstance(api_keys, list):
            self.api_keys = [k.strip() for k in api_keys if isinstance(k, str) and k.strip()]
        else:
            self.api_keys = []

        # Cố định model gemini-3.5-flash theo yêu cầu
        self.model_name = 'gemini-3.5-flash'

    def extract_invoice_data(self, file_path: str, expected_tags=None):
        if not self.api_keys:
            return {"error": "Chưa tìm thấy danh sách Gemini API Keys trong file .env."}

        file_extension = file_path.split('.')[-1].lower()
        if file_extension not in ['png', 'jpg', 'jpeg', 'pdf', 'docx', 'xlsx']:
            return {"error": f"Định dạng file .{file_extension} chưa được hỗ trợ."}

        # Chuẩn bị trước nội dung cho ảnh / word / excel để không phải parse nhiều lần khi đổi key
        image_content = None
        text_content_payload = None

        try:
            if file_extension in ['png', 'jpg', 'jpeg']:
                image_content = Image.open(file_path)
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
        - NGUYÊN TẮC: Đọc kỹ phần thông tin chung (header/footer) của tài liệu. Trích xuất CHÍNH XÁC nội dung CÓ THẬT trên tài liệu (Tên công ty, Đại diện, Địa chỉ, MST, Điện thoại, Số tài khoản, Ngày tháng...).
        - KỶ LUẬT THÉP: TUYỆT ĐỐI KHÔNG tự suy luận, không bịa đặt. Thông tin nào hoàn toàn không có trên tài liệu, BẮT BUỘC để chuỗi rỗng `""` để người dùng tự nhập tay.
        """ if expected_tags else ""

        prompt = f"""
        Bạn là Kế toán trưởng và Chuyên gia ERP. Nhiệm vụ của bạn là đọc và BÓC TÁCH CHÍNH XÁC thông tin từ tài liệu theo đúng nguyên tắc "Chứng từ có sao ghi vậy".
        {tags_instruction}

        YÊU CẦU NGHIÊM NGẶT (STRICT EXTRACTION):
        1. Trả về JSON hợp lệ. Tiền tệ, số lượng là SỐ (không phẩy, không khoảng trắng, không "VNĐ").
        2. CÓ GÌ GHI NẤY: Trích xuất y hệt thông tin trên giấy. 
           - Trạng thái các cột: Tuyệt đối KHÔNG TỰ ĐOÁN mục đích sử dụng ("muc_dich_su_dung"). Nếu không ghi cột mục đích, bắt buộc để chuỗi rỗng "".
           - Với trường "ghi_chu": Chỉ ghi nhận nếu tài liệu có cột Ghi chú, hoặc có các cột phụ (như "Xuất xứ", "Quy cách") thì gộp chung vào. Nếu không có, để trống "".
        3. XÁC ĐỊNH THUẾ VAT: Đọc kỹ tài liệu xem giá trên bảng kê/báo giá là ĐÃ BAO GỒM VAT hay CHƯA BAO GỒM VAT. Ghi vào `thong_tin_vat.da_bao_gom_vat` (true/false) và trích xuất `thue_suat` (ví dụ 8, 10, 5, 0. Mặc định 8 nếu không ghi rõ).
        4. Tính toán lại số liệu, ghi lỗi vào "danh_sach_canh_bao". Nếu chuẩn 100%, để mảng rỗng [].

        CẤU TRÚC JSON BẮT BUỘC:
        {{
          "thong_tin_nha_cung_cap": {{"ten_cong_ty": "", "dia_chi": "", "dien_thoai": "", "ma_so_thue": "", "email": ""}},
          "thong_tin_chung": {{"loai_chung_tu": "", "so_chung_tu": "", "ngay_thang_nam": ""}},
          "thong_tin_khach_hang": {{"ten_khach_hang": "", "dia_chi": "", "ma_so_thue": ""}},
          "danh_sach_hang_hoa": [{{"stt": 1, "ten_hang_hoa": "", "don_vi_tinh": "", "so_luong": 0, "don_gia": 0, "thanh_tien": 0, "muc_dich_su_dung": "", "ghi_chu": ""}}],
          "tong_ket_tien": {{"tong_tien_truoc_thue": 0, "thue_suat_vat": "", "tien_thue_vat": 0, "tong_tien_thanh_toan": 0, "so_tien_viet_bang_chu": ""}},
          "thong_tin_vat": {{"da_bao_gom_vat": false, "thue_suat": 8}},
          "thong_tin_dong": {{}},
          "danh_sach_canh_bao": []
        }}
        """

        num_keys = len(self.api_keys)
        start_index = AIExtractor._current_key_index
        last_error_msg = ""
        hit_quota_error = False

        # Thử lần lượt qua toàn bộ các key trong cụm (Round-Robin)
        for offset in range(num_keys):
            key_idx = (start_index + offset) % num_keys
            api_key = self.api_keys[key_idx]
            client = None
            uploaded_pdf = None

            try:
                client = genai.Client(api_key=api_key)

                # Chuẩn bị media_content cho client tương ứng
                if file_extension in ['png', 'jpg', 'jpeg']:
                    media_content = image_content
                elif file_extension == 'pdf':
                    uploaded_pdf = client.files.upload(file=file_path)
                    media_content = uploaded_pdf
                else:
                    media_content = text_content_payload

                # Gọi AI với model gemini-3.5-flash
                response = client.models.generate_content(
                    model=self.model_name,
                    contents=[media_content, prompt],
                    config={"response_mime_type": "application/json"}
                )

                if not response or not response.text:
                    continue

                text_result = response.text.strip()
                if text_result.startswith("```json"): 
                    text_result = text_result[7:-3].strip()
                elif text_result.startswith("```"): 
                    text_result = text_result[3:-3].strip()

                parsed_data = json.loads(text_result)

                # Thành công: Cập nhật vị trí xoay vòng cho lần gọi kế tiếp
                AIExtractor._current_key_index = (key_idx + 1) % num_keys
                return parsed_data

            except json.JSONDecodeError:
                last_error_msg = "AI trả về định dạng dữ liệu không hợp lệ. Đang chuyển máy chủ khác thử lại..."
                continue
            except Exception as e:
                err_str = str(e)
                err_lower = err_str.lower()
                last_error_msg = err_str

                # Nhận diện lỗi 429 / Quota / Resource Exhausted
                if "429" in err_str or "resource_exhausted" in err_lower or "quota" in err_lower:
                    hit_quota_error = True
                    print(f"[AIExtractor] Key #{key_idx + 1}/{num_keys} đạt giới hạn quota ({self.model_name}). Đang chuyển sang key tiếp theo...")
                    continue
                else:
                    # Lỗi mạng tạm thời hoặc kết nối, thử key kế tiếp
                    continue
            finally:
                # Luôn xóa file PDF tạm trên server Google sau khi xử lý
                if uploaded_pdf and client:
                    try:
                        client.files.delete(name=uploaded_pdf.name)
                    except Exception:
                        pass

        # Khi toàn bộ 3 key đều hết quota / quá tải
        if hit_quota_error:
            return {
                "error": "⚠️ Toàn bộ 3 máy chủ AI hiện đang đạt giới hạn yêu cầu tạm thời. "
                         "Vui lòng đợi trong ít phút (khoảng 1 - 2 phút) để 3 máy chủ tự động phục hồi và hoạt động lại bình thường."
            }
        
        return {"error": f"Không thể trích xuất tài liệu lúc này. Chi tiết: {last_error_msg}"}