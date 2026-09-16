import os
import sys
import subprocess

if __name__ == "__main__":
    # Đảm bảo in tiếng Việt chuẩn trên terminal Windows
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
        except Exception:
            pass

    print("🚀 Đang khởi động Hệ thống Kế toán AI...")

    # Tự động phát hiện và ưu tiên sử dụng python trong thư mục venv nếu tồn tại
    app_dir = os.path.dirname(os.path.abspath(__file__))
    venv_python = os.path.join(app_dir, "venv", "Scripts", "python.exe")
    
    python_bin = venv_python if os.path.exists(venv_python) else sys.executable

    # Chạy Streamlit
    cmd = [python_bin, "-m", "streamlit", "run", "app.py"]
    print(f"🔧 Môi trường Python: {python_bin}")
    
    try:
        subprocess.run(cmd, cwd=app_dir)
    except KeyboardInterrupt:
        print("\n🛑 Đã dừng ứng dụng.")