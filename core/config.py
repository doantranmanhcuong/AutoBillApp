import os
import re
from dotenv import load_dotenv

# Tự động nạp các biến từ file .env (override=True để luôn cập nhật mới nhất khi chạy local)
load_dotenv(override=True)

API_KEYS = []

# 1. Nạp từ biến môi trường (chạy local qua .env)
env_multi = os.getenv("GEMINI_API_KEYS", "")
env_single = os.getenv("GEMINI_API_KEY", "")

if env_multi:
    for k in re.split(r'[,;\n\r]+', env_multi):
        k_clean = k.strip().strip("'").strip('"')
        if k_clean and k_clean not in API_KEYS:
            API_KEYS.append(k_clean)

if env_single:
    s_clean = env_single.strip().strip("'").strip('"')
    if s_clean and s_clean not in API_KEYS:
        API_KEYS.append(s_clean)

# 2. Hỗ trợ tự động nạp Streamlit Cloud Secrets khi triển khai online
try:
    import streamlit as st
    if hasattr(st, "secrets"):
        # Trường hợp GEMINI_API_KEYS là danh sách list hoặc chuỗi string
        if "GEMINI_API_KEYS" in st.secrets:
            val = st.secrets["GEMINI_API_KEYS"]
            if isinstance(val, (list, tuple)):
                for k in val:
                    k_str = str(k).strip()
                    if k_str and k_str not in API_KEYS:
                        API_KEYS.append(k_str)
            elif isinstance(val, str):
                for k in val.split(","):
                    k_str = k.strip()
                    if k_str and k_str not in API_KEYS:
                        API_KEYS.append(k_str)

        if "GEMINI_API_KEY" in st.secrets:
            val_single = str(st.secrets["GEMINI_API_KEY"]).strip()
            if val_single and val_single not in API_KEYS:
                API_KEYS.append(val_single)
except Exception:
    pass

# Key mặc định đầu tiên để đảm bảo tương thích ngược
API_KEY = API_KEYS[0] if API_KEYS else None

# Xác định đường dẫn thư mục gốc và thư mục templates an toàn tuyệt đối
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

if not os.path.exists(TEMPLATE_DIR):
    os.makedirs(TEMPLATE_DIR)