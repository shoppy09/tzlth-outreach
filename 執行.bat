@echo off
chcp 65001 >nul
set PYTHONUTF8=1
python main.py
if %errorlevel% neq 0 (
    echo.
    echo 程式發生錯誤，請截圖回報。
    pause
)
