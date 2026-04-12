#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
職涯顧問合作拓展系統 — 網頁版
Flask Web App
"""
import sys, os
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8","utf8"):
    sys.stdout.reconfigure(encoding="utf-8")

from flask import Flask, render_template, request, jsonify, redirect, url_for, send_file
import pandas as pd
from datetime import datetime, timedelta
import threading, webbrowser, json

# 引入原有模組
from main import (
    load_t, load_l, _save, load_tmpl, fill_tmpl, send_email,
    _subject, _mark_sent, _gen_pdf, TYPES, STATUSES,
    CONTACT_MODES, PRIORITIES, ROUNDS, ensure_dirs, init_data,
    LINE_TMPL, PHONE_TMPL, TEMPLATES_DIR, OUTPUTS_DIR, DATA_FILE,
    TYPE_TMPL
)

app = Flask(__name__)
app.secret_key = "tzlth_outreach_2026"

# 啟動時自動初始化（支援 gunicorn / Render 雲端部署）
ensure_dirs()
init_data()

@app.context_processor
def inject_now():
    return {"now": datetime.now().strftime("%Y-%m-%d %H:%M")}

BASE_DIR = Path(__file__).parent.parent

# ── 輔助函數 ──────────────────────────────────────────────────────────────────
def df_to_records(df):
    df = df.copy()
    # 字串欄位：NaN → '' （避免 Jinja2 顯示 "nan"）
    obj_cols = df.select_dtypes(include='object').columns
    df[obj_cols] = df[obj_cols].fillna('')
    # 數值欄位：NaN → None
    for col in df.select_dtypes(exclude='object').columns:
        df[col] = df[col].where(df[col].notna(), other=None)
    return df.to_dict(orient="records")

def get_stats():
    df_t = load_t()
    df_l = load_l()
    today = pd.Timestamp(datetime.now().date())
    week_later = today + pd.Timedelta(days=7)
    df_t["下次跟進日期"] = pd.to_datetime(df_t["下次跟進日期"], errors="coerce")
    active = df_t[~df_t["狀態"].isin(["合作中", "暫不合適"])]
    overdue = active[active["下次跟進日期"].notna() & (active["下次跟進日期"] <= today)]
    due_soon = active[
        active["下次跟進日期"].notna() &
        (active["下次跟進日期"] > today) &
        (active["下次跟進日期"] <= week_later)
    ]
    return {
        "total":    int(len(df_t)),
        "unsent":   int((df_t["狀態"] == "未聯絡").sum()),
        "sent":     int((df_t["狀態"] == "已發信").sum()),
        "replied":  int((df_t["狀態"] == "已回覆").sum()),
        "meeting":  int((df_t["狀態"] == "已安排會議").sum()),
        "deal":     int((df_t["狀態"] == "合作中").sum()),
        "overdue":  int(len(overdue)),
        "due_soon": int(len(due_soon)),
        "overdue_list":  df_to_records(overdue[["ID","單位名稱","狀態","下次跟進日期"]].head(5)),
        "due_soon_list": df_to_records(due_soon[["ID","單位名稱","狀態","下次跟進日期"]].head(5)),
        "today_list":    df_to_records(active[
            active["下次跟進日期"].notna() &
            (active["下次跟進日期"] == today)
        ][["ID","單位名稱","狀態","下次跟進日期"]]),
        "recent_log":    df_to_records(df_l.tail(8)[["日期","單位名稱","聯絡方式","結果","內容摘要"]]) if not df_l.empty else [],
        "type_stats": {t: int((df_t["單位類型"]==t).sum()) for t in TYPES},
        "status_stats": {s: int((df_t["狀態"]==s).sum()) for s in STATUSES},
    }

# ── 路由：儀表板 ──────────────────────────────────────────────────────────────
@app.route("/")
def index():
    stats = get_stats()
    return render_template("index.html", stats=stats)

# ── 路由：目標單位 ────────────────────────────────────────────────────────────
@app.route("/targets")
def targets():
    df_t = load_t()
    fs = request.args.get("status", "")
    ft = request.args.get("type", "")
    fp = request.args.get("priority", "")
    if fs: df_t = df_t[df_t["狀態"] == fs]
    if ft: df_t = df_t[df_t["單位類型"] == ft]
    if fp: df_t = df_t[df_t["優先序"] == fp]
    records = df_to_records(df_t)
    return render_template("targets.html",
        targets=records, types=TYPES, statuses=STATUSES,
        priorities=PRIORITIES, contact_modes=CONTACT_MODES,
        fs=fs, ft=ft, fp=fp)

@app.route("/targets/add", methods=["POST"])
def add_target():
    df_t = load_t(); df_l = load_l()
    nid = int(df_t["ID"].max()) + 1 if not df_t.empty else 1
    row = {
        "ID": nid,
        "單位名稱": request.form["name"],
        "單位類型": request.form["type"],
        "聯絡人":   request.form.get("contact",""),
        "職稱":     request.form.get("title",""),
        "Email":    request.form.get("email",""),
        "電話":     request.form.get("phone",""),
        "地區":     request.form.get("region",""),
        "狀態":     "未聯絡",
        "優先序":   request.form.get("priority","中"),
        "聯絡方式": request.form.get("contact_mode","Email"),
        "輪次":     "第一輪",
        "發信日期": None, "最後跟進日期": None, "下次跟進日期": None,
        "備註":     request.form.get("note",""),
    }
    df_t = pd.concat([df_t, pd.DataFrame([row])], ignore_index=True)
    _save(df_t, df_l)
    return redirect(url_for("targets"))

@app.route("/targets/<int:tid>/update-status", methods=["POST"])
def update_status(tid):
    df_t = load_t(); df_l = load_l()
    new_s = request.form["status"]
    note  = request.form.get("note","")
    today = datetime.now().strftime("%Y-%m-%d")
    rows  = df_t[df_t["ID"]==tid]
    if rows.empty: return jsonify({"ok": False})
    r = rows.iloc[0]
    # 先把日期欄轉成 object 避免 dtype 衝突
    for col in ["最後跟進日期", "下次跟進日期", "發信日期"]:
        if col in df_t.columns:
            df_t[col] = df_t[col].astype(object)
    df_t.loc[df_t["ID"]==tid, "狀態"] = new_s
    df_t.loc[df_t["ID"]==tid, "最後跟進日期"] = today
    if new_s == "已回覆":
        df_t.loc[df_t["ID"]==tid, "輪次"] = "第二輪"
        df_t.loc[df_t["ID"]==tid, "下次跟進日期"] = (datetime.now()+timedelta(days=3)).strftime("%Y-%m-%d")
    elif new_s == "已發信":
        df_t.loc[df_t["ID"]==tid, "下次跟進日期"] = (datetime.now()+timedelta(days=7)).strftime("%Y-%m-%d")
    if note: df_t.loc[df_t["ID"]==tid, "備註"] = note
    log = {"日期":today,"單位ID":tid,"單位名稱":r["單位名稱"],
           "聯絡方式":"其他","內容摘要":note or "狀態更新","結果":new_s}
    df_l = pd.concat([df_l, pd.DataFrame([log])], ignore_index=True)
    _save(df_t, df_l)
    return jsonify({"ok": True, "new_status": new_s,
                    "upgraded": new_s=="已回覆"})

@app.route("/targets/<int:tid>/update-email", methods=["POST"])
def update_email(tid):
    df_t = load_t(); df_l = load_l()
    new_email = request.form.get("email","").strip()
    df_t.loc[df_t["ID"]==tid, "Email"] = new_email
    _save(df_t, df_l)
    return jsonify({"ok": True})

# ── 路由：發信 ────────────────────────────────────────────────────────────────
@app.route("/compose/<int:tid>")
def compose(tid):
    df_t = load_t()
    rows = df_t[df_t["ID"]==tid]
    if rows.empty: return redirect(url_for("targets"))
    r = rows.iloc[0]
    round_ = str(r.get("輪次","第一輪"))
    tmpl = load_tmpl(str(r["單位類型"]), round_)
    body = fill_tmpl(tmpl, r) if tmpl else "(找不到對應模板)"
    return render_template("compose.html", target=r.to_dict(), body=body,
                           rounds=ROUNDS, statuses=STATUSES)

@app.route("/compose/<int:tid>/preview", methods=["POST"])
def preview(tid):
    df_t = load_t()
    rows = df_t[df_t["ID"]==tid]
    if rows.empty: return jsonify({"body":""})
    r = rows.iloc[0]
    round_ = request.form.get("round", str(r.get("輪次","第一輪")))
    tmpl = load_tmpl(str(r["單位類型"]), round_)
    body = fill_tmpl(tmpl, r) if tmpl else "(找不到模板)"
    subj = _subject(r)
    return jsonify({"body": body, "subject": subj})

@app.route("/compose/<int:tid>/send", methods=["POST"])
def send(tid):
    df_t = load_t(); df_l = load_l()
    rows = df_t[df_t["ID"]==tid]
    if rows.empty: return jsonify({"ok":False,"msg":"找不到單位"})
    r = rows.iloc[0]
    body    = request.form.get("body","")
    subj    = request.form.get("subject","合作提案｜蒲朝棟")
    to_addr = str(r["Email"]) if pd.notna(r["Email"]) else ""
    gen_pdf = request.form.get("gen_pdf","0") == "1"

    pdf_path = None
    if gen_pdf:
        pdf_path = _gen_pdf(r)

    if not to_addr:
        return jsonify({"ok":False,"msg":"此單位尚無 Email，請先填入"})

    sent = send_email(to_addr, subj, body, pdf_path, quiet=True)
    if sent:
        df_t, df_l = _mark_sent(df_t, df_l, tid, r)
        _save(df_t, df_l)
        fname = f"{datetime.now().strftime('%Y%m%d')}_{r['單位名稱']}.txt"
        (OUTPUTS_DIR/fname).write_text(body, encoding="utf-8")
        return jsonify({"ok":True,"msg":f"✓ 已發送至 {to_addr}"})
    else:
        return jsonify({"ok":False,"msg":"發送失敗，請確認 config.env 設定"})

@app.route("/compose/<int:tid>/save", methods=["POST"])
def save_draft(tid):
    df_t = load_t(); df_l = load_l()
    rows = df_t[df_t["ID"]==tid]
    if rows.empty: return jsonify({"ok":False})
    r = rows.iloc[0]
    body = request.form.get("body","")
    fname = f"{datetime.now().strftime('%Y%m%d')}_{r['單位名稱']}.txt"
    (OUTPUTS_DIR/fname).write_text(body, encoding="utf-8")
    df_t, df_l = _mark_sent(df_t, df_l, tid, r)
    _save(df_t, df_l)
    return jsonify({"ok":True,"msg":f"✓ 已儲存至 outputs/{fname}"})

# ── 路由：腳本 ────────────────────────────────────────────────────────────────
@app.route("/script/<string:stype>/<int:tid>")
def get_script(stype, tid):
    df_t = load_t()
    rows = df_t[df_t["ID"]==tid]
    if rows.empty: return jsonify({"text":""})
    r = rows.iloc[0]
    tmpl_map = LINE_TMPL if stype=="line" else PHONE_TMPL
    fname = tmpl_map.get(str(r["單位類型"]))
    if not fname: return jsonify({"text":"找不到對應腳本"})
    path = TEMPLATES_DIR / fname
    if not path.exists(): return jsonify({"text":"檔案不存在"})
    contact = str(r["聯絡人"]) if pd.notna(r["聯絡人"]) and str(r["聯絡人"]).strip() else ""
    text = (path.read_text(encoding="utf-8")
            .replace("{{單位名稱}}", str(r["單位名稱"]))
            .replace("{{聯絡人}}", contact or "您")
            .replace("{{地區}}", str(r["地區"]) if pd.notna(r["地區"]) else ""))
    return jsonify({"text": text, "name": str(r["單位名稱"])})

# ── 路由：批次發信 ────────────────────────────────────────────────────────────
@app.route("/batch")
def batch():
    df_t = load_t()
    unsent = df_t[df_t["狀態"]=="未聯絡"]
    return render_template("batch.html", targets=df_to_records(unsent),
                           types=TYPES, priorities=PRIORITIES)

@app.route("/batch/send", methods=["POST"])
def batch_send():
    df_t = load_t(); df_l = load_l()
    ids      = request.json.get("ids", [])
    gen_pdf  = request.json.get("gen_pdf", False)
    do_send  = request.json.get("send", False)
    results  = []
    for tid in ids:
        tid = int(tid)
        rows = df_t[df_t["ID"]==tid]
        if rows.empty: continue
        r = rows.iloc[0]
        round_ = str(r.get("輪次","第一輪"))
        tmpl = load_tmpl(str(r["單位類型"]), round_)
        if not tmpl:
            results.append({"id":tid,"name":str(r["單位名稱"]),"status":"error","msg":"找不到模板"})
            continue
        body = fill_tmpl(tmpl, r)
        fname = f"{datetime.now().strftime('%Y%m%d')}_{r['單位名稱']}_{round_}.txt"
        (OUTPUTS_DIR/fname).write_text(body, encoding="utf-8")
        pdf_path = _gen_pdf(r) if gen_pdf else None
        email_str = str(r["Email"]) if pd.notna(r["Email"]) and str(r["Email"]).strip() else ""
        sent = False
        if do_send and email_str:
            sent = send_email(email_str, _subject(r), body, pdf_path, quiet=True)
        df_t, df_l = _mark_sent(df_t, df_l, tid, r)
        results.append({
            "id": tid, "name": str(r["單位名稱"]),
            "status": "sent" if sent else "saved",
            "msg": f"發送至 {email_str}" if sent else "已存檔（無Email或未啟用發送）"
        })
    _save(df_t, df_l)
    return jsonify({"ok":True,"results":results})

# ── 路由：聯絡記錄 ────────────────────────────────────────────────────────────
@app.route("/records")
def records():
    df_l = load_l()
    recs = df_to_records(df_l.tail(50)) if not df_l.empty else []
    recs.reverse()
    return render_template("records.html", records=recs)

# ── 路由：週報 ────────────────────────────────────────────────────────────────
@app.route("/report")
def report():
    stats = get_stats()
    return render_template("report.html", stats=stats)

@app.route("/report/generate")
def generate_report():
    try:
        sys.path.insert(0, str(BASE_DIR))
        from reporter import generate_report as gen
        path = gen(open_after=False)
        return send_file(str(path), as_attachment=True,
                         download_name=path.name, mimetype="application/pdf")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ── 路由：PDF 提案書 ──────────────────────────────────────────────────────────
@app.route("/pdf/<int:tid>")
def gen_pdf_route(tid):
    df_t = load_t()
    rows = df_t[df_t["ID"]==tid]
    if rows.empty: return jsonify({"error":"找不到單位"}), 404
    r = rows.iloc[0]
    pdf_path = _gen_pdf(r)
    if pdf_path:
        return send_file(str(pdf_path), as_attachment=True,
                         download_name=pdf_path.name, mimetype="application/pdf")
    return jsonify({"error":"PDF 生成失敗"}), 500

# ── 路由：回信偵測 ────────────────────────────────────────────────────────────
@app.route("/check-replies", methods=["POST"])
def check_replies_route():
    try:
        from scheduler import check_replies
        updated = check_replies(quiet=True)
        return jsonify({"ok":True,"updated":len(updated),
                        "names":[n for _,n in updated]})
    except Exception as e:
        return jsonify({"ok":False,"msg":str(e)})

# ── 路由：看板 ────────────────────────────────────────────────────────────────
@app.route("/kanban")
def kanban():
    df_t = load_t()
    records = df_to_records(df_t)
    return render_template("kanban.html", targets=records, statuses=STATUSES)

# ── 路由：單位詳情 ────────────────────────────────────────────────────────────
@app.route("/target/<int:tid>")
def target_detail(tid):
    df_t = load_t()
    df_l = load_l()
    rows = df_t[df_t["ID"] == tid]
    if rows.empty:
        return redirect(url_for("targets"))
    target = rows.iloc[0].to_dict()
    # 此單位的聯絡記錄（含單位ID欄位比對，或用名稱）
    if "單位ID" in df_l.columns:
        logs = df_l[df_l["單位ID"] == tid]
    else:
        logs = df_l[df_l["單位名稱"] == target["單位名稱"]]
    logs = df_to_records(logs.sort_values("日期", ascending=False))
    return render_template("target_detail.html",
        target=target, logs=logs,
        statuses=STATUSES, priorities=PRIORITIES,
        contact_modes=CONTACT_MODES)

# ── 路由：新增聯絡記錄 ────────────────────────────────────────────────────────
@app.route("/targets/<int:tid>/add-log", methods=["POST"])
def add_log(tid):
    df_t = load_t()
    df_l = load_l()
    rows = df_t[df_t["ID"] == tid]
    if rows.empty:
        return jsonify({"ok": False, "msg": "找不到單位"})
    r = rows.iloc[0]
    today = datetime.now().strftime("%Y-%m-%d")
    mode    = request.form.get("mode", "其他")
    result  = request.form.get("result", "")
    summary = request.form.get("summary", "")
    log = {"日期": today, "單位ID": tid, "單位名稱": r["單位名稱"],
           "聯絡方式": mode, "內容摘要": summary, "結果": result}
    df_l = pd.concat([df_l, pd.DataFrame([log])], ignore_index=True)
    # 同步更新目標單位的狀態
    if result and result != r["狀態"]:
        for col in ["狀態", "最後跟進日期"]:
            df_t[col] = df_t[col].astype(object)
        df_t.loc[df_t["ID"] == tid, "狀態"] = result
        df_t.loc[df_t["ID"] == tid, "最後跟進日期"] = today
    _save(df_t, df_l)
    return jsonify({"ok": True, "log": log})

# ── 路由：API 趨勢資料 ────────────────────────────────────────────────────────
@app.route("/api/trend")
def api_trend():
    """回傳最近 8 週的聯絡數量趨勢"""
    df_l = load_l()
    weeks = []
    for i in range(7, -1, -1):
        start = datetime.now() - timedelta(days=(i+1)*7)
        end   = datetime.now() - timedelta(days=i*7)
        label = start.strftime("%m/%d")
        if df_l.empty:
            count = 0
        else:
            df_l["日期"] = pd.to_datetime(df_l["日期"], errors="coerce")
            count = int(((df_l["日期"] >= pd.Timestamp(start)) &
                         (df_l["日期"] < pd.Timestamp(end))).sum())
        weeks.append({"label": label, "count": count})
    return jsonify(weeks)

# ── 路由：API A/B 測試 ────────────────────────────────────────────────────────
@app.route("/api/ab-test")
def api_ab_test():
    """比較奇偶 ID 的回覆率"""
    df_t = load_t()
    replied = ["已回覆", "已安排會議", "合作中"]
    sent_all = df_t[df_t["狀態"].isin(["已發信"] + replied)]

    def calc(group):
        total = len(group)
        r = group["狀態"].isin(replied).sum()
        return {"total": int(total), "replied": int(r),
                "rate": round(r / total * 100, 1) if total > 0 else 0}

    odd  = sent_all[sent_all["ID"] % 2 == 1]
    even = sent_all[sent_all["ID"] % 2 == 0]
    return jsonify({"A": calc(odd), "B": calc(even)})

# ── 路由：範本管理 ────────────────────────────────────────────────────────────
@app.route("/templates")
def tmpl_editor():
    tmpl_list = []
    type_short = {
        "退輔會/榮服處":"退輔會", "政府勞動局處":"政府勞動局",
        "大專院校":"大專院校",   "社區大學":"社區大學",
        "企業HR":"企業HR",       "人資協會":"人資協會"
    }
    for (typ, rnd), fname in TYPE_TMPL.items():
        path = TEMPLATES_DIR / fname
        content = path.read_text(encoding="utf-8") if path.exists() else ""
        key = fname.replace(".txt","")
        tmpl_list.append({
            "key": key, "fname": fname,
            "label": f"{type_short.get(typ,typ)} {rnd}",
            "type": typ, "round": rnd, "content": content
        })
    # LINE & 電話話術
    for prefix, label_pfx in [("LINE_","LINE"), ("電話話術_","電話話術")]:
        for typ in TYPES:
            short = type_short.get(typ, typ)
            fname = f"{prefix}{short}.txt"
            path  = TEMPLATES_DIR / fname
            if not path.exists():
                # 嘗試完整類型名
                fname = f"{prefix}{typ}.txt"
                path  = TEMPLATES_DIR / fname
            content = path.read_text(encoding="utf-8") if path.exists() else ""
            key = fname.replace(".txt","")
            tmpl_list.append({
                "key": key, "fname": fname,
                "label": f"{label_pfx} — {short}",
                "type": typ, "round": label_pfx, "content": content
            })
    return render_template("tmpl_editor.html", templates=tmpl_list, types=TYPES)

@app.route("/api/templates/<path:key>/save", methods=["POST"])
def save_template(key):
    try:
        content = request.json.get("content","")
        fname = key + ".txt"
        path  = TEMPLATES_DIR / fname
        path.write_text(content, encoding="utf-8")
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)})

@app.route("/api/templates/<path:key>/reset")
def reset_template(key):
    """返回目前磁碟上的版本（不做任何修改）"""
    fname = key + ".txt"
    path  = TEMPLATES_DIR / fname
    if path.exists():
        return jsonify({"ok": True, "content": path.read_text(encoding="utf-8")})
    return jsonify({"ok": False, "msg": "找不到範本"})

# ── 路由：匯出 Excel ───────────────────────────────────────────────────────────
@app.route("/export")
def export_page():
    stats = get_stats()
    return render_template("export.html", stats=stats)

@app.route("/export/targets")
def export_targets():
    import io
    df_t = load_t()
    filt = request.args.get("filter","")
    if filt == "contacted":
        df_t = df_t[df_t["狀態"] != "未聯絡"]
    elif filt == "unsent":
        df_t = df_t[df_t["狀態"] == "未聯絡"]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_t.to_excel(w, index=False, sheet_name="目標單位")
    buf.seek(0)
    fname = f"目標單位_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export/logs")
def export_logs():
    import io
    df_l = load_l()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_l.to_excel(w, index=False, sheet_name="聯絡記錄")
    buf.seek(0)
    fname = f"聯絡記錄_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export/backup/<filename>")
def export_backup(filename):
    backup_dir = BASE_DIR / "data" / "backup"
    path = backup_dir / filename
    if not path.exists():
        return "找不到備份檔", 404
    return send_file(str(path), as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── 路由：立即備份 ────────────────────────────────────────────────────────────
@app.route("/api/backup-now", methods=["POST"])
def backup_now():
    try:
        import shutil
        backup_dir = BASE_DIR / "data" / "backup"
        backup_dir.mkdir(exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out   = backup_dir / f"targets_{stamp}.xlsx"
        shutil.copy2(DATA_FILE, out)
        return jsonify({"ok": True, "file": out.name})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)})

# ── 路由：Gmail 草稿 ──────────────────────────────────────────────────────────
@app.route("/compose/<int:tid>/save-draft-gmail", methods=["POST"])
def save_gmail_draft(tid):
    """存入 Gmail 草稿匣（使用 SMTP 不支援，改用 IMAP APPEND）"""
    try:
        import imaplib, email as emaillib
        from email.mime.text    import MIMEText
        from email.mime.multipart import MIMEMultipart

        # 讀取設定
        cfg = {}
        cfg_path = BASE_DIR / "config.env"
        if cfg_path.exists():
            for line in cfg_path.read_text(encoding="utf-8").splitlines():
                if "=" in line and not line.startswith("#"):
                    k,v = line.split("=",1)
                    cfg[k.strip()] = v.strip()

        sender  = cfg.get("SENDER_EMAIL","")
        pwd     = cfg.get("SENDER_PASSWORD","")
        if not sender or not pwd:
            return jsonify({"ok":False,"msg":"請先在 config.env 設定 Gmail 帳號密碼"})

        df_t = load_t()
        rows = df_t[df_t["ID"]==tid]
        if rows.empty:
            return jsonify({"ok":False,"msg":"找不到單位"})
        r = rows.iloc[0]

        subject = request.form.get("subject","")
        body    = request.form.get("body","")
        to_addr = str(r["Email"]) if pd.notna(r["Email"]) else ""

        msg = MIMEMultipart()
        msg["From"]    = sender
        msg["To"]      = to_addr
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # IMAP APPEND 到 [Gmail]/Drafts
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(sender, pwd)
        imap.append("[Gmail]/Drafts", "\\Draft",
                    imaplib.Time2Internaldate(datetime.now()),
                    msg.as_bytes())
        imap.logout()
        return jsonify({"ok":True,"msg":"已存入 Gmail 草稿匣"})
    except Exception as e:
        return jsonify({"ok":False,"msg":f"存入草稿失敗：{e}"})

# ── 路由：網路自動補全 ────────────────────────────────────────────────────────
@app.route("/targets/<int:tid>/autofill", methods=["POST"])
def autofill_target(tid):
    """嘗試從網路搜尋補全 Email / 聯絡人資訊"""
    try:
        import urllib.request, urllib.parse, re, html

        df_t = load_t()
        df_l = load_l()
        rows = df_t[df_t["ID"]==tid]
        if rows.empty:
            return jsonify({"ok":False,"msg":"找不到單位"})
        r = rows.iloc[0]
        name = str(r["單位名稱"])

        # 用 Google 搜尋取得聯絡資訊
        query = urllib.parse.quote(f"{name} email 聯絡")
        url   = f"https://www.google.com/search?q={query}"
        req   = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        })
        with urllib.request.urlopen(req, timeout=8) as resp:
            raw = resp.read().decode("utf-8", errors="ignore")

        # 從頁面提取 email
        emails = list(set(re.findall(
            r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', raw
        )))
        # 過濾掉 google / w3 等無關 email
        skip = {"google","w3.org","schema","example","sentry","cloudflare"}
        emails = [e for e in emails if not any(s in e for s in skip)][:5]

        found = {}
        if emails:
            found["suggested_emails"] = emails

        return jsonify({"ok":True, "found": found, "name": name})
    except Exception as e:
        return jsonify({"ok":False, "msg":f"搜尋失敗：{e}"})

# ── 路由：API 通知（今日待辦 JSON）────────────────────────────────────────────
@app.route("/api/today-todos")
def api_today_todos():
    """回傳今日需跟進的單位，給前端瀏覽器通知使用"""
    df_t = load_t()
    today = pd.Timestamp(datetime.now().date())
    df_t["下次跟進日期"] = pd.to_datetime(df_t["下次跟進日期"], errors="coerce")
    active = df_t[~df_t["狀態"].isin(["合作中","暫不合適"])]
    overdue = active[active["下次跟進日期"].notna() & (active["下次跟進日期"] < today)]
    today_due = active[active["下次跟進日期"].notna() & (active["下次跟進日期"] == today)]
    return jsonify({
        "overdue": df_to_records(overdue[["ID","單位名稱","狀態"]]),
        "today":   df_to_records(today_due[["ID","單位名稱","狀態"]]),
        "count":   int(len(overdue)) + int(len(today_due))
    })

# ── 路由：API ─────────────────────────────────────────────────────────────────
@app.route("/api/stats")
def api_stats():
    return jsonify(get_stats())

@app.route("/api/targets")
def api_targets():
    df_t = load_t()
    return jsonify(df_to_records(df_t))

# ── 路由：備份管理 ────────────────────────────────────────────────────────────
@app.route("/api/backups")
def api_backups():
    backup_dir = BASE_DIR / "data" / "backup"
    if not backup_dir.exists():
        return jsonify([])
    files = sorted(backup_dir.glob("targets_*.xlsx"), reverse=True)
    return jsonify([{"name": f.name, "size": f.stat().st_size,
                     "date": f.stem.replace("targets_","")} for f in files[:10]])

# ── 啟動 ──────────────────────────────────────────────────────────────────────
def open_browser():
    import time; time.sleep(1.2)
    webbrowser.open("http://localhost:5000")

if __name__ == "__main__":
    ensure_dirs(); init_data()
    print("=" * 50)
    print("  職涯顧問合作拓展系統 — 網頁版")
    print("  http://localhost:5000")
    print("  按 Ctrl+C 關閉")
    print("=" * 50)
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(debug=False, port=5000, use_reloader=False)
