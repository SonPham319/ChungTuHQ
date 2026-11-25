import sys
import os

# Configure UTF-8 for console output on Windows
if sys.platform == 'win32':
    if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
        try:
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass
    if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
        try:
            sys.stderr.reconfigure(encoding='utf-8', errors='replace')
        except Exception:
            pass

# Add current dir to python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui import AppGUI

def main():
    print("Khởi chạy ứng dụng Tạo BK & Mẫu 25 NQ...")
    app = AppGUI()
    app.mainloop()

if __name__ == '__main__':
    main()

