@echo off
chcp 65001 >nul
echo 正在設定每日提醒工作排程...
echo.

REM 設定每天早上 8:30 執行提醒
set TASK_NAME=職涯合作系統每日提醒
set SCRIPT_PATH=%~dp0每日提醒.py
set PYTHON_PATH=python

REM 刪除舊排程（若有）
schtasks /delete /tn "%TASK_NAME%" /f >nul 2>&1

REM 建立新排程
schtasks /create /tn "%TASK_NAME%" /tr "%PYTHON_PATH% \"%SCRIPT_PATH%\"" /sc daily /st 08:30 /f

if %errorlevel%==0 (
    echo ✓ 設定成功！每天早上 8:30 會自動跳出提醒
    echo.
    echo 工作名稱：%TASK_NAME%
    echo 執行時間：每天 08:30
    echo 執行程式：%SCRIPT_PATH%
) else (
    echo ✗ 設定失敗，請以「系統管理員身份」執行此檔案
)

echo.
pause
