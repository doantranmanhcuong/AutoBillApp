from google import genai
from PIL import Image
import json
import pandas as pd
import docx

class AIExtractor:
    def __init__(self, api_key: str):
        self.client = genai.Client(api_key=api_key)
        self.model_name = 'gemini-3.5-flash' 

    def extract_invoice_data(self, file_path: str, expected_tags=None):
        uploaded_pdf = None # Biến tạm để lưu file PDF trên server AI
        try:
            file_extension = file_path.split('.')[-1].lower()
            
            if file_extension in ['png', 'jpg', 'jpeg']:
                media_content = Image.open(file_path)
            elif file_extension == 'pdf':
                # Upload thẳng PDF lên Gemini để giữ nguyên cấu trúc (đọc được cả PDF scan)
                uploaded_pdf = self.client.files.upload(file=file_path)
                media_content = uploaded_pdf
            elif file_extension == 'docx':
                doc = docx.Document(file_path)
                text_content = "\n".join([p.text for p in doc.paragraphs if p.text.strip() != ""])
                media_content = f"Dữ liệu văn bản bóc tách từ file Word:\n{text_content}"
            elif file_extension == 'xlsx':
                df_dict = pd.read_excel(file_path, sheet_name=None)
                text_content = ""
                for sheet_name, df in df_dict.items():
                    text_content += f"\n--- Sheet: {sheet_name} ---\n{df.to_string(index=False)}"
                media_content = f"Dữ liệu bảng tính bóc tách từ file Excel:\n{text_content}"
            else:
                return {"error": f"Định dạng file .{file_extension} chưa được hỗ trợ."}
            
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
            3. Tính toán lại số liệu, ghi lỗi vào "danh_sach_canh_bao". Nếu chuẩn 100%, để mảng rỗng [].

            CẤU TRÚC JSON BẮT BUỘC:
            {{
              "thong_tin_nha_cung_cap": {{"ten_cong_ty": "", "dia_chi": "", "dien_thoai": "", "ma_so_thue": "", "email": ""}},
              "thong_tin_chung": {{"loai_chung_tu": "", "so_chung_tu": "", "ngay_thang_nam": ""}},
              "thong_tin_khach_hang": {{"ten_khach_hang": "", "dia_chi": "", "ma_so_thue": ""}},
              "danh_sach_hang_hoa": [{{"stt": 1, "ten_hang_hoa": "", "don_vi_tinh": "", "so_luong": 0, "don_gia": 0, "thanh_tien": 0, "muc_dich_su_dung": "", "ghi_chu": ""}}],
              "tong_ket_tien": {{"tong_tien_truoc_thue": 0, "thue_suat_vat": "", "tien_thue_vat": 0, "tong_tien_thanh_toan": 0, "so_tien_viet_bang_chu": ""}},
              "thong_tin_dong": {{}},
              "danh_sach_canh_bao": []
            }}
            """

            response = self.client.models.generate_content(
                model=self.model_name,
                contents=[media_content, prompt],
                config={"response_mime_type": "application/json"}
            )
            
            text_result = response.text.strip()
            if text_result.startswith("```json"): text_result = text_result[7:-3].strip()
            elif text_result.startswith("```"): text_result = text_result[3:-3].strip()
                
            return json.loads(text_result)
            
        except json.JSONDecodeError:
            return {"error": "AI trả về định dạng dữ liệu không hợp lệ. Vui lòng thử lại."}
        except Exception as e:
            return {"error": f"Lỗi hệ thống trong quá trình trích xuất: {str(e)}"}
        finally:
            # Dọn dẹp: Luôn xóa file PDF tạm trên máy chủ Google AI sau khi xử lý xong (dù thành công hay lỗi)
            if uploaded_pdf:
                try:
                    self.client.files.delete(name=uploaded_pdf.name)
                except Exception:
                    pass