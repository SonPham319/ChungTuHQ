@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist "C:\Users\PC\AppData\Local\Programs\Python\Python312\python.exe" (
    "C:\Users\PC\AppData\Local\Programs\Python\Python312\python.exe" main.py
) else (
    python main.py
)
pause

