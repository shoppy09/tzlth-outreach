@echo off
chcp 65001 >nul
title 合作拓展系統 — 網頁版
set PYTHONUTF8=1

echo.
echo  ╔══════════════════════════════════╗
echo  ║   職涯顧問合作拓展系統 — 網頁版    ║
echo  ║   蒲朝棟 Tim Pu                   ║
echo  ╚══════════════════════════════════╝
echo.
echo  啟動中，請稍候...
echo  瀏覽器將自動開啟 http://localhost:5000
echo.
echo  關閉此視窗即可停止系統
echo.

cd /d "%~dp0"
python web\app.py

pause
