# 📑 AutoBillApp - Hệ Thống Lập Hồ Sơ & Chứng Từ Kế Toán Tự Động

**AutoBillApp** là ứng dụng tự động hóa quy trình lập chứng từ kế toán, đề nghị thanh toán và hợp đồng kinh tế từ hóa đơn/báo giá đầu vào bằng việc kết hợp **Streamlit** và **Google Gemini AI**.

---

## 🌟 Tính Năng Nổi Bật

- 📥 **Hỗ trợ đa định dạng:** Đọc file PDF (kể cả PDF scan đa trang), Ảnh (PNG, JPG, JPEG), Word (`.docx`) và Excel (`.xlsx`).
- 📑 **Gộp nhiều hóa đơn cùng một Nhà cung cấp (Multi-Invoice Consolidation):**
  - Cho phép tải lên nhiều hóa đơn cùng lúc từ một đối tác/nhà cung cấp.
  - Tự động đối chiếu tính đồng nhất của Nhà cung cấp (cảnh báo nếu phát hiện sai lệch MST hoặc tên công ty).
  - Tự động ghép nối toàn bộ danh mục hàng hóa, đánh lại STT liên tục, gắn tag nguồn gốc chứng từ `[HĐ: ...]`, và tổng hợp số chứng từ.
- ⚡ **Hệ thống Multi-Key Pool & Round-Robin:**
  - Hỗ trợ xoay vòng cụm API Keys liên tục giữa các yêu cầu.
  - Tự động failover khi một key chạm hạn mức rate-limit (`429 RESOURCE_EXHAUSTED`), đảm bảo hệ thống hoạt động ổn định và không chập chờn.
- 🤖 **Bóc tách thông minh với Gemini 3.5 Flash:** Trích xuất chính xác thông tin nhà cung cấp, khách hàng, số chứng từ, ngày tháng và bảng kê hàng hóa theo nguyên tắc nghiêm ngặt *"Chứng từ có sao ghi vậy"*.
- ⚖️ **Linh hoạt xử lý Thuế VAT (2 chế độ):**
  - **Báo giá CHƯA gồm VAT:** Tự động cộng tiền thuế vào tổng thanh toán.
  - **Báo giá ĐÃ bao gồm VAT:** Bóc tách doanh thu trước thuế và tiền thuế từ tổng thanh toán.
  - Tự động dịch số tiền thanh toán ra chữ tiếng Việt chuẩn ngữ pháp kế toán.
- 📄 **Bảo toàn 100% định dạng Word (`.docx`):** Thuật toán *Run-Span Replacement* giữ nguyên vẹn Font chữ, Cỡ chữ, Màu sắc, In đậm, In nghiêng của mẫu gốc.
- 📊 **Xử lý linh hoạt mẫu Excel (`.xlsx`):** Tự động nhân bản dòng, copy định dạng (borders, fill, font, alignment), tính toán ô gộp (`merged cells`) và dọn dẹp sạch tag rỗng.
- 📦 **Đóng gói trọn bộ:** Nén tất cả biểu mẫu đã điền thành 1 file ZIP tải về ngay lập tức.
- 💬 **Mẫu tin nhắn Zalo 1-click:** Sinh sẵn tin nhắn Zalo kèm tổng tiền chuẩn định dạng để gửi lãnh đạo duyệt nhanh.

---

## 🚀 Hướng Dẫn Cài Đặt & Sử Dụng

### 1. Yêu cầu hệ thống
- Python 3.10 trở lên.
- Google Gemini API Key (Lấy miễn phí tại [Google AI Studio](https://aistudio.google.com/)).

### 2. Cài đặt môi trường Local
```bash
# Clone dự án về máy
git clone https://github.com/doantranmanhcuong/AutoBillApp.git
cd AutoBillApp

# Tạo môi trường ảo (khuyến nghị)
python -m venv venv

# Kích hoạt môi trường ảo (Windows)
.\venv\Scripts\activate

# Cài đặt các thư viện cần thiết
pip install -r requirements.txt
```

### 3. Cấu hình API Key
Tạo file `.env` tại thư mục gốc của dự án (tham khảo file `.env.example`):
```env
# Điền danh sách các key cách nhau bởi dấu phẩy để hệ thống tự động xoay vòng
GEMINI_API_KEYS=key1,key2,key3
GEMINI_API_KEY=key1
```

### 4. Khởi động ứng dụng Local
```bash
python main.py
```
Hoặc:
```bash
streamlit run app.py
```
Mở trình duyệt truy cập: `http://localhost:8501` (hoặc `http://<IP_Mạng_LAN>:8501` để dùng chung trong văn phòng).

---

## ☁️ Hướng Dẫn Triển Khai Lên Streamlit Cloud

1. Đẩy mã nguồn dự án lên GitHub cá nhân của bạn.
2. Truy cập [share.streamlit.io](https://share.streamlit.io) và chọn kho lưu trữ GitHub của bạn.
3. Cấu hình triển khai:
   - **Main file path:** `app.py`
4. Cấu hình Secrets (trong mục **Advanced settings** -> **Secrets** của ứng dụng Streamlit Cloud):
   ```toml
   GEMINI_API_KEYS = "AIzaSyA_key1,AIzaSyB_key2,AIzaSyC_key3"
   GEMINI_API_KEY = "AIzaSyA_key1"
   ```
5. Bấm **Deploy!** Ứng dụng sẽ tự động cài đặt các dependencies từ `requirements.txt` và nạp key từ Secrets an toàn.

---

## 📂 Cấu Trúc Thư Mục

```text
AutoBillApp/
├── core/
│   ├── ai_extractor.py         # Module kết nối và trích xuất dữ liệu bằng Gemini AI xoay vòng key
│   ├── invoice_merger.py       # Module kiểm tra và gộp nhiều hóa đơn cùng một Nhà cung cấp
│   ├── config.py               # Quản lý cấu hình tập trung, biến môi trường và Streamlit Secrets
│   ├── document_builder.py     # Engine điền dữ liệu vào biểu mẫu Word và Excel
│   └── utils/
│       └── helpers.py          # Đọc tiền thành chữ tiếng Việt & Quét thẻ template
├── templates/                  # Thư mục chứa các biểu mẫu kế toán (.docx, .xlsx)
├── ui/
│   └── components.py           # Giao diện văn phòng, thẻ tài chính & mẫu tin nhắn Zalo
├── .env.example                # File mẫu cấu hình biến môi trường
├── .gitignore                  # Cấu hình bỏ qua file bí mật và file tạm
├── app.py                      # Ứng dụng chính Streamlit
├── main.py                     # Entrypoint khởi chạy ứng dụng local
└── requirements.txt            # Danh sách các thư viện Python
```

---

## 🔒 Bảo Mật & An Toàn Dữ Liệu
- Toàn bộ file chứng từ tạm được xử lý trong bộ nhớ RAM hoặc thư mục tạm tự hủy (`tempfile.TemporaryDirectory`).
- Khóa bí mật API Key và các file tạm được bảo vệ trong `.gitignore`, tuyệt đối không bị đẩy lên kho lưu trữ GitHub.
