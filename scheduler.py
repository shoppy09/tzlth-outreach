#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
排程發信 + IMAP 回信自動偵測
蒲朝棟 Tim Pu
"""
import sys, imaplib, email, os, json, smtplib
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8","utf8"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

BASE_DIR   = Path(__file__).parent
DATA_FILE  = BASE_DIR / "data" / "targets.xlsx"
CONFIG_FILE= BASE_DIR / "config.env"
QUEUE_FILE = BASE_DIR / "data" / "send_queue.json"
OUTPUTS_DIR= BASE_DIR / "outputs"

# ── 顏色 ─────────────────────────────────────────────────────────────────────
G="\033[92m"; Y="\033[93m"; R="\033[91m"; C="\033[96m"; BOLD="\033[1m"; Z="\033[0m"

def load_cfg():
    cfg = {}
    if CONFIG_FILE.exists():
        for line in CONFIG_FILE.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k,v = line.split("=",1); cfg[k.strip()] = v.strip()
    return cfg

def load_t(): return pd.read_excel(DATA_FILE, sheet_name="目標單位", dtype={"ID":int})
def load_l(): return pd.read_excel(DATA_FILE, sheet_name="聯絡記錄")
def save_all(df_t, df_l):
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
        df_t.to_excel(w, sheet_name="目標單位", index=False)
        df_l.to_excel(w, sheet_name="聯絡記錄", index=False)

# ══════════════════════════════════════════════════════════════════
# IMAP 回信偵測
# ══════════════════════════════════════════════════════════════════
def decode_str(s):
    """解碼 Email 標頭"""
    parts = decode_header(s)
    result = []
    for p, enc in parts:
        if isinstance(p, bytes):
            result.append(p.decode(enc or "utf-8", errors="ignore"))
        else:
            result.append(p)
    return "".join(result)

def check_replies(quiet=False):
    """
    連線 Gmail IMAP，檢查是否有目標單位的回信。
    若有，自動將狀態更新為「已回覆」。
    """
    cfg = load_cfg()
    email_addr = cfg.get("SENDER_EMAIL","")
    password   = cfg.get("SENDER_PASSWORD","")
    if not email_addr or not password:
        print(f"{Y}⚠ 請先在 config.env 設定 Email 與密碼{Z}")
        return []

    df_t = load_t()
    df_l = load_l()

    # 建立目標 Email 對照表（target email → 單位 ID + 名稱）
    target_emails = {}
    for _, r in df_t.iterrows():
        if pd.notna(r["Email"]) and str(r["Email"]).strip():
            domain = str(r["Email"]).strip().lower()
            target_emails[domain] = (int(r["ID"]), str(r["單位名稱"]))

    updated = []
    try:
        if not quiet: print(f"\n{C}正在連線 Gmail IMAP...{Z}")
        with imaplib.IMAP4_SSL("imap.gmail.com") as mail:
            mail.login(email_addr, password)
            mail.select("inbox")

            # 搜尋未讀信件
            _, msg_ids = mail.search(None, "UNSEEN")
            ids = msg_ids[0].split()
            if not quiet: print(f"  發現 {len(ids)} 封未讀信件，檢查中...")

            for mid in ids:
                _, data = mail.fetch(mid, "(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                from_addr = decode_str(msg.get("From","")).lower()
                subject   = decode_str(msg.get("Subject",""))
                date_str  = msg.get("Date","")

                # 比對寄件者是否在目標清單
                matched_id   = None
                matched_name = None
                for target_email, (tid, tname) in target_emails.items():
                    if target_email in from_addr:
                        # 確認對方狀態是「已發信」
                        row = df_t[df_t["ID"]==tid]
                        if not row.empty and str(row.iloc[0]["狀態"]) in ("已發信","未聯絡"):
                            matched_id   = tid
                            matched_name = tname
                            break

                if matched_id:
                    today = datetime.now().strftime("%Y-%m-%d")
                    follow= (datetime.now()+timedelta(days=3)).strftime("%Y-%m-%d")
                    df_t.loc[df_t["ID"]==matched_id, "狀態"]          = "已回覆"
                    df_t.loc[df_t["ID"]==matched_id, "輪次"]           = "第二輪"
                    df_t.loc[df_t["ID"]==matched_id, "最後跟進日期"]   = today
                    df_t.loc[df_t["ID"]==matched_id, "下次跟進日期"]   = follow

                    log = {"日期":today,"單位ID":matched_id,"單位名稱":matched_name,
                           "聯絡方式":"Email回覆","內容摘要":f"對方回信：{subject[:30]}","結果":"已回覆"}
                    df_l = pd.concat([df_l,pd.DataFrame([log])],ignore_index=True)
                    updated.append((matched_id, matched_name))
                    print(f"  {G}✓ 偵測到回覆：ID {matched_id} {matched_name}{Z}")

        if updated:
            save_all(df_t, df_l)
            print(f"\n{G}✓ 已自動更新 {len(updated)} 筆狀態為「已回覆」，並升級至第二輪{Z}")
        else:
            if not quiet: print(f"  {Y}目前無目標單位的新回信{Z}")

    except Exception as e:
        print(f"{R}✗ IMAP 連線失敗：{e}{Z}")
        print(f"  請確認 config.env 的帳號密碼，並確認已開啟「應用程式密碼」")

    return updated


# ══════════════════════════════════════════════════════════════════
# 排程發信佇列
# ══════════════════════════════════════════════════════════════════
def load_queue():
    if not QUEUE_FILE.exists(): return []
    return json.loads(QUEUE_FILE.read_text(encoding="utf-8"))

def save_queue(q):
    QUEUE_FILE.parent.mkdir(exist_ok=True)
    QUEUE_FILE.write_text(json.dumps(q, ensure_ascii=False, indent=2), encoding="utf-8")

def add_to_queue(target_ids: list, send_at: str):
    """
    加入排程佇列
    target_ids: list of int
    send_at: "YYYY-MM-DD HH:MM"
    """
    q = load_queue()
    q.append({
        "ids":     target_ids,
        "send_at": send_at,
        "status":  "pending",
        "added":   datetime.now().strftime("%Y-%m-%d %H:%M"),
    })
    save_queue(q)
    print(f"{G}✓ 已加入排程：{len(target_ids)} 筆，預定發送時間：{send_at}{Z}")

def process_queue():
    """檢查佇列，發送到時間的排程"""
    from main import load_tmpl, fill_tmpl, send_email, _subject, _mark_sent, _gen_pdf
    q   = load_queue()
    now = datetime.now()
    updated_q = []
    any_sent  = False

    for item in q:
        if item["status"] != "pending":
            updated_q.append(item); continue
        send_at = datetime.strptime(item["send_at"], "%Y-%m-%d %H:%M")
        if now < send_at:
            updated_q.append(item); continue

        # 到時間了，發送
        print(f"\n{C}執行排程發信 (排定時間: {item['send_at']}){Z}")
        df_t = load_t(); df_l = load_l()
        for tid in item["ids"]:
            rows = df_t[df_t["ID"]==tid]
            if rows.empty: continue
            r     = rows.iloc[0]
            round_= str(r.get("輪次","第一輪"))
            tmpl  = load_tmpl(str(r["單位類型"]), round_)
            if not tmpl: continue
            body  = fill_tmpl(tmpl, r)
            fname = f"{now.strftime('%Y%m%d')}_{r['單位名稱']}_{round_}.txt"
            (OUTPUTS_DIR/fname).write_text(body, encoding="utf-8")
            email_str = str(r["Email"]) if pd.notna(r["Email"]) else ""
            if email_str.strip():
                send_email(email_str, _subject(r), body, quiet=True)
                print(f"  {G}✓ 已發送：{r['單位名稱']}{Z}")
            df_t, df_l = _mark_sent(df_t, df_l, tid, r)
        save_all(df_t, df_l)
        item["status"] = "done"
        item["done_at"]= now.strftime("%Y-%m-%d %H:%M")
        any_sent = True
        updated_q.append(item)

    save_queue(updated_q)
    return any_sent

def view_queue():
    q = load_queue()
    if not q: print(f"  {Y}排程佇列目前為空{Z}"); return
    print(f"\n  {'狀態':<8}  {'預定時間':<18}  {'單位數':>5}  {'加入時間'}")
    print("  " + "─"*55)
    for item in q:
        st = f"{G}完成{Z}" if item["status"]=="done" else f"{Y}待發{Z}"
        print(f"  {st}      {item['send_at']:<18}  {len(item['ids']):>5}  {item['added']}")

def schedule_interactive():
    """互動式排程設定"""
    print("\n  最佳發信時機建議：週二或週三 09:00–10:00")
    print("  格式範例：2026-04-15 09:00\n")

    df_t = load_t()
    # 顯示未聯絡的
    pending = df_t[df_t["狀態"]=="未聯絡"]
    if pending.empty:
        print(f"  {Y}目前無待發信的單位{Z}"); return

    print(f"  目前有 {len(pending)} 筆未聯絡單位，可加入排程")
    print(f"  請輸入目標 ID（逗號分隔，如 1,3,5）或輸入 all 全選：")
    ids_input = input("  ID: ").strip()
    if ids_input.lower()=="all":
        ids = list(pending["ID"].astype(int))
    else:
        ids = [int(x.strip()) for x in ids_input.split(",") if x.strip().isdigit()]

    send_at = input("  預定發送時間 (YYYY-MM-DD HH:MM): ").strip()
    try:
        datetime.strptime(send_at, "%Y-%m-%d %H:%M")
    except ValueError:
        print(f"  {R}時間格式錯誤{Z}"); return

    add_to_queue(ids, send_at)


if __name__ == "__main__":
    import sys
    if len(sys.argv)>1 and sys.argv[1]=="check":
        check_replies()
    elif len(sys.argv)>1 and sys.argv[1]=="run":
        process_queue()
    else:
        print("用法：python scheduler.py check   # 檢查回信")
        print("      python scheduler.py run     # 執行排程佇列")
