import os
from dotenv import load_dotenv

# Tự động tìm và nạp các biến từ file .env
load_dotenv()

API_KEY = os.getenv("GEMINI_API_KEY")
TEMPLATE_DIR = "templates"

# Đảm bảo thư mục templates luôn tồn tại
if not os.path.exists(TEMPLATE_DIR):
    os.makedirs(TEMPLATE_DIR)