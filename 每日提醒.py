#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日提醒程式（由 Windows 工作排程器呼叫）
每天早上 8:30 執行，跳出待跟進提醒
"""
import sys, os, subprocess
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8","utf8"):
    sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

BASE_DIR  = Path(__file__).parent
DATA_FILE = BASE_DIR / "data" / "targets.xlsx"

def get_summary():
    if not DATA_FILE.exists():
        return "資料庫尚未建立", []

    df_t = pd.read_excel(DATA_FILE, sheet_name="目標單位")
    df_t["下次跟進日期"] = pd.to_datetime(df_t["下次跟進日期"], errors="coerce")
    today = datetime.now().date()

    active = df_t[~df_t["狀態"].isin(["合作中","暫不合適"])]
    overdue  = active[active["下次跟進日期"].notna() & (active["下次跟進日期"].dt.date < today)]
    due_today= active[active["下次跟進日期"].notna() & (active["下次跟進日期"].dt.date == today)]
    this_week= active[active["下次跟進日期"].notna() &
                      (active["下次跟進日期"].dt.date > today) &
                      (active["下次跟進日期"].dt.date <= today + timedelta(days=7))]

    lines = []
    if not overdue.empty:
        lines.append(f"⚠ 逾期跟進：{len(overdue)} 筆")
        for _,r in overdue.iterrows():
            lines.append(f"  • ID{int(r['ID'])} {r['單位名稱']} ({r['狀態']})")
    if not due_today.empty:
        lines.append(f"📅 今日跟進：{len(due_today)} 筆")
        for _,r in due_today.iterrows():
            lines.append(f"  • ID{int(r['ID'])} {r['單位名稱']} ({r['狀態']})")
    if not this_week.empty:
        lines.append(f"📋 本週跟進：{len(this_week)} 筆")

    total_unsent = (df_t["狀態"]=="未聯絡").sum()
    total_deal   = (df_t["狀態"]=="合作中").sum()
    lines.append(f"\n待聯絡 {total_unsent} 筆 | 合作中 {total_deal} 筆")

    title = "🔔 職涯合作系統提醒"
    if len(overdue) > 0:
        title = f"⚠️ 有 {len(overdue)} 筆逾期！職涯合作系統提醒"
    elif len(due_today) > 0:
        title = f"📅 今日有 {len(due_today)} 筆需跟進｜職涯合作系統"

    return title, lines

def show_notification(title, message):
    """使用 Windows 通知"""
    try:
        import ctypes
        # 使用 powershell 顯示 Toast 通知
        ps_script = f"""
Add-Type -AssemblyName System.Windows.Forms
$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = [System.Drawing.SystemIcons]::Information
$notify.Visible = $true
$notify.ShowBalloonTip(10000, '{title}', '{message}', [System.Windows.Forms.ToolTipIcon]::Info)
Start-Sleep -Seconds 10
$notify.Dispose()
"""
        subprocess.run(["powershell", "-Command", ps_script],
                       capture_output=True, timeout=15)
    except Exception:
        pass

def show_popup(title, lines):
    """使用 MessageBox 顯示提醒"""
    message = "\n".join(lines) if lines else "今日無待跟進項目，繼續保持！"
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x40 | 0x1000)
    except Exception:
        # fallback: 直接打開系統
        pass

def main():
    title, lines = get_summary()
    if not lines:
        lines = ["今日無待跟進項目，繼續保持！"]

    # 顯示彈窗
    show_popup(title, lines)

    # 問是否要開啟系統
    try:
        import ctypes
        open_q = "是否立即開啟合作拓展系統？"
        result = ctypes.windll.user32.MessageBoxW(0, open_q, "職涯合作系統", 0x04 | 0x1000)
        if result == 6:  # Yes
            bat = BASE_DIR / "執行.bat"
            if bat.exists():
                os.startfile(str(bat))
    except Exception:
        pass

if __name__ == "__main__":
    main()
