#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
週報自動生成模組
蒲朝棟 Tim Pu
"""
import sys
if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8","utf8"):
    sys.stdout.reconfigure(encoding="utf-8")

import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

BASE_DIR  = Path(__file__).parent
DATA_FILE = BASE_DIR / "data" / "targets.xlsx"
REPORTS_DIR = BASE_DIR / "reports"

STATUSES = ["未聯絡","已發信","已回覆","已安排會議","合作中","暫不合適"]
TYPES    = ["退輔會/榮服處","政府勞動局處","大專院校","社區大學","企業HR","人資協會"]

NAVY  = colors.HexColor("#1B3A6B")
TEAL  = colors.HexColor("#2A7D8C")
GREEN = colors.HexColor("#2E7D32")
RED   = colors.HexColor("#C62828")
AMBER = colors.HexColor("#F57F17")
LIGHT = colors.HexColor("#F0F5FA")
GREY  = colors.HexColor("#555555")
WHITE = colors.white

def _add_mappings():
    """Declare MSJhei / MSJheiBd as ReportLab font families."""
    from reportlab.lib.fonts import addMapping
    addMapping("MSJhei",   0, 0, "MSJhei")
    addMapping("MSJhei",   1, 0, "MSJheiBd")
    addMapping("MSJhei",   0, 1, "MSJhei")
    addMapping("MSJhei",   1, 1, "MSJheiBd")
    addMapping("MSJheiBd", 0, 0, "MSJheiBd")
    addMapping("MSJheiBd", 1, 0, "MSJheiBd")
    addMapping("MSJheiBd", 0, 1, "MSJheiBd")
    addMapping("MSJheiBd", 1, 1, "MSJheiBd")

def _register_pair(regular_path, bold_path, subfont_index=None):
    """Register MSJhei/MSJheiBd font pair (OTF/TTF or TTC with index)."""
    kwargs = {}
    if subfont_index is not None:
        kwargs["subfontIndex"] = subfont_index
    pdfmetrics.registerFont(TTFont("MSJhei",   regular_path, **kwargs))
    pdfmetrics.registerFont(TTFont("MSJheiBd", bold_path,   **kwargs))
    _add_mappings()

def register_fonts():
    import glob
    # Already registered in this process — skip
    if "MSJhei" in pdfmetrics._fonts:
        return

    # OTF/TTF single-font candidates (no subfontIndex needed)
    # Debian python:3.11-slim + fonts-noto-cjk → /usr/share/fonts/opentype/noto/
    otf_candidates = [
        ("C:/Windows/Fonts/msjh.ttc",    "C:/Windows/Fonts/msjhbd.ttc"),
        ("/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf",
         "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Bold.otf"),
        ("/usr/share/fonts/opentype/noto/NotoSansCJKtc-Regular.otf",
         "/usr/share/fonts/opentype/noto/NotoSansCJKtc-Bold.otf"),
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.otf",
         "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.otf"),
        ("/usr/share/fonts/truetype/noto/NotoSansCJKsc-Regular.otf",
         "/usr/share/fonts/truetype/noto/NotoSansCJKsc-Bold.otf"),
        ("/usr/share/fonts/truetype/noto/NotoSansCJKtc-Regular.otf",
         "/usr/share/fonts/truetype/noto/NotoSansCJKtc-Bold.otf"),
        ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
         "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"),
    ]
    for regular, bold in otf_candidates:
        if Path(regular).exists():
            bold_path = bold if Path(bold).exists() else regular
            try:
                _register_pair(regular, bold_path)
                print(f"[fonts] OK (otf): {regular}", file=sys.stderr)
                return
            except Exception as exc:
                print(f"[fonts] FAIL {regular}: {exc}", file=sys.stderr)

    # TTC collection candidates — require subfontIndex=0
    ttc_candidates = [
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
         "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
        ("/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
         "/usr/share/fonts/noto-cjk/NotoSansCJK-Bold.ttc"),
    ]
    for regular, bold in ttc_candidates:
        if Path(regular).exists():
            bold_path = bold if Path(bold).exists() else regular
            try:
                _register_pair(regular, bold_path, subfont_index=0)
                print(f"[fonts] OK (ttc): {regular}", file=sys.stderr)
                return
            except Exception as exc:
                print(f"[fonts] FAIL TTC {regular}: {exc}", file=sys.stderr)

    # Glob fallback — try OTF first, then TTC
    for pattern in ["/usr/share/fonts/**/*[Cc][Jj][Kk]*[Ss][Cc]-Regular*.otf",
                    "/usr/share/fonts/**/*[Cc][Jj][Kk]*Regular*.otf",
                    "/usr/share/fonts/**/*[Cc][Jj][Kk]*Regular*.ttf",
                    "/usr/share/fonts/**/*[Cc][Jj][Kk]*.ttc"]:
        matches = sorted(glob.glob(pattern, recursive=True))
        print(f"[fonts] glob {pattern}: {matches}", file=sys.stderr)
        if matches:
            m = matches[0]
            try:
                if m.lower().endswith(".ttc"):
                    _register_pair(m, m, subfont_index=0)
                else:
                    _register_pair(m, m)
                print(f"[fonts] OK (glob): {m}", file=sys.stderr)
                return
            except Exception as exc:
                print(f"[fonts] FAIL glob {m}: {exc}", file=sys.stderr)

    print("[fonts] WARNING: no CJK font registered — PDF will crash", file=sys.stderr)

def make_styles():
    return {
        "title":   ParagraphStyle("title",   fontName="MSJheiBd", fontSize=18, textColor=WHITE,   leading=24),
        "sub":     ParagraphStyle("sub",     fontName="MSJhei",   fontSize=10, textColor=colors.HexColor("#CCE0FF"), leading=14),
        "section": ParagraphStyle("section", fontName="MSJheiBd", fontSize=12, textColor=NAVY,    spaceBefore=12, spaceAfter=6, leading=16),
        "body":    ParagraphStyle("body",    fontName="MSJhei",   fontSize=9,  textColor=GREY,    spaceAfter=3,  leading=14),
        "num":     ParagraphStyle("num",     fontName="MSJheiBd", fontSize=22, textColor=NAVY,    leading=28, alignment=1),
        "numlbl":  ParagraphStyle("numlbl",  fontName="MSJhei",   fontSize=8,  textColor=GREY,    leading=12, alignment=1),
        "good":    ParagraphStyle("good",    fontName="MSJheiBd", fontSize=9,  textColor=GREEN,   leading=14),
        "warn":    ParagraphStyle("warn",    fontName="MSJheiBd", fontSize=9,  textColor=AMBER,   leading=14),
        "bad":     ParagraphStyle("bad",     fontName="MSJheiBd", fontSize=9,  textColor=RED,     leading=14),
        "tag":     ParagraphStyle("tag",     fontName="MSJheiBd", fontSize=8,  textColor=TEAL,    leading=12),
        "footer":  ParagraphStyle("footer",  fontName="MSJhei",   fontSize=7,  textColor=colors.HexColor("#999"), leading=11),
    }

def load_data():
    df_t = pd.read_excel(DATA_FILE, sheet_name="目標單位", dtype={"ID":int})
    df_l = pd.read_excel(DATA_FILE, sheet_name="聯絡記錄")
    return df_t, df_l

def generate_report(open_after=True) -> Path:
    register_fonts()
    REPORTS_DIR.mkdir(exist_ok=True)
    styles = make_styles()
    today  = datetime.now()
    week_ago = today - timedelta(days=7)

    df_t, df_l = load_data()
    df_l["日期"] = pd.to_datetime(df_l["日期"], errors="coerce")
    this_week_log = df_l[df_l["日期"] >= week_ago]

    fname = f"週報_{today.strftime('%Y%m%d')}.pdf"
    out   = REPORTS_DIR / fname
    W, H  = A4
    margin = 1.8*cm

    doc = SimpleDocTemplate(str(out), pagesize=A4,
                            leftMargin=margin, rightMargin=margin,
                            topMargin=1.2*cm, bottomMargin=1.5*cm)
    story = []

    # ── Banner ───────────────────────────────────────────────────────────────
    banner = Table([[
        Paragraph("合作拓展週報", styles["title"]),
        Paragraph(f"{today.strftime('%Y 年 %m 月 %d 日')}<br/>蒲朝棟 Tim Pu", styles["sub"]),
    ]], colWidths=[(W-2*margin)*0.65, (W-2*margin)*0.35])
    banner.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), NAVY),
        ("LEFTPADDING",   (0,0),(-1,-1), 14),
        ("RIGHTPADDING",  (0,0),(-1,-1), 14),
        ("TOPPADDING",    (0,0),(-1,-1), 12),
        ("BOTTOMPADDING", (0,0),(-1,-1), 12),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("ALIGN",         (1,0),(1,-1),  "RIGHT"),
    ]))
    story.extend([banner, Spacer(1, 0.3*cm)])

    # ── 關鍵數字 ─────────────────────────────────────────────────────────────
    story.append(Paragraph("▍ 整體進度", styles["section"]))
    total   = len(df_t)
    sent    = (df_t["狀態"]=="已發信").sum()
    replied = (df_t["狀態"]=="已回覆").sum()
    meeting = (df_t["狀態"]=="已安排會議").sum()
    deal    = (df_t["狀態"]=="合作中").sum()
    unsent  = (df_t["狀態"]=="未聯絡").sum()
    reply_rate = f"{replied/max(sent+replied+meeting+deal,1)*100:.0f}%"

    kpi_data = [[
        Paragraph(str(total),   styles["num"]),
        Paragraph(str(unsent),  styles["num"]),
        Paragraph(str(sent),    styles["num"]),
        Paragraph(str(replied), styles["num"]),
        Paragraph(str(meeting), styles["num"]),
        Paragraph(str(deal),    styles["num"]),
    ],[
        Paragraph("目標單位總數", styles["numlbl"]),
        Paragraph("待聯絡",      styles["numlbl"]),
        Paragraph("已發信",      styles["numlbl"]),
        Paragraph("已回覆",      styles["numlbl"]),
        Paragraph("已安排會議",  styles["numlbl"]),
        Paragraph("合作中",      styles["numlbl"]),
    ]]
    cw = (W-2*margin)/6
    kpi_table = Table(kpi_data, colWidths=[cw]*6, rowHeights=[36, 18])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), LIGHT),
        ("BACKGROUND",    (5,0),(5,-1),  colors.HexColor("#E8F5E9")),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("ALIGN",         (0,0),(-1,-1), "CENTER"),
        ("GRID",          (0,0),(-1,-1), 0.5, colors.HexColor("#DDDDDD")),
    ]))
    story.extend([kpi_table, Spacer(1, 0.2*cm)])

    # ── 本週活動 ─────────────────────────────────────────────────────────────
    story.append(Paragraph("▍ 本週聯絡紀錄", styles["section"]))
    if this_week_log.empty:
        story.append(Paragraph("本週尚無聯絡記錄", styles["body"]))
    else:
        log_data = [[
            Paragraph("日期", styles["tag"]),
            Paragraph("單位名稱", styles["tag"]),
            Paragraph("聯絡方式", styles["tag"]),
            Paragraph("結果", styles["tag"]),
            Paragraph("內容摘要", styles["tag"]),
        ]]
        for _, r in this_week_log.iterrows():
            log_data.append([
                Paragraph(str(r["日期"])[:10], styles["body"]),
                Paragraph(str(r["單位名稱"])[:18], styles["body"]),
                Paragraph(str(r["聯絡方式"]), styles["body"]),
                Paragraph(str(r["結果"]), styles["body"]),
                Paragraph(str(r["內容摘要"])[:30], styles["body"]),
            ])
        cws = [(W-2*margin)*x for x in [0.12, 0.22, 0.12, 0.16, 0.38]]
        log_table = Table(log_data, colWidths=cws)
        log_table.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,0),  TEAL),
            ("TEXTCOLOR",     (0,0),(-1,0),  WHITE),
            ("ROWBACKGROUNDS",(0,1),(-1,-1), [LIGHT, colors.HexColor("#E8EFF8")]),
            ("LEFTPADDING",   (0,0),(-1,-1), 6),
            ("RIGHTPADDING",  (0,0),(-1,-1), 6),
            ("TOPPADDING",    (0,0),(-1,-1), 4),
            ("BOTTOMPADDING", (0,0),(-1,-1), 4),
            ("VALIGN",        (0,0),(-1,-1), "TOP"),
        ]))
        story.append(log_table)
    story.append(Spacer(1, 0.2*cm))

    # ── 各類型進度 ───────────────────────────────────────────────────────────
    story.append(Paragraph("▍ 各類型進度", styles["section"]))
    type_data = [[
        Paragraph("類型", styles["tag"]),
        Paragraph("總數", styles["tag"]),
        Paragraph("未聯絡", styles["tag"]),
        Paragraph("已發信", styles["tag"]),
        Paragraph("已回覆+", styles["tag"]),
        Paragraph("進度", styles["tag"]),
    ]]
    for t in TYPES:
        sub = df_t[df_t["單位類型"]==t]
        n   = len(sub)
        ns  = (sub["狀態"]=="未聯絡").sum()
        nf  = (sub["狀態"]=="已發信").sum()
        nr  = sub["狀態"].isin(["已回覆","已安排會議","合作中"]).sum()
        prog= f"{(n-ns)/max(n,1)*100:.0f}%"
        type_data.append([
            Paragraph(t, styles["body"]),
            Paragraph(str(n),  styles["body"]),
            Paragraph(str(ns), styles["warn"] if ns>0 else styles["body"]),
            Paragraph(str(nf), styles["body"]),
            Paragraph(str(nr), styles["good"] if nr>0 else styles["body"]),
            Paragraph(prog,    styles["good"] if int(prog[:-1])>50 else styles["body"]),
        ])
    cws2 = [(W-2*margin)*x for x in [0.28,0.1,0.14,0.14,0.14,0.2]]
    type_table = Table(type_data, colWidths=cws2)
    type_table.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0),  TEAL),
        ("TEXTCOLOR",     (0,0),(-1,0),  WHITE),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [LIGHT, colors.HexColor("#E8EFF8")]),
        ("LEFTPADDING",   (0,0),(-1,-1), 8),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("ALIGN",         (1,0),(-1,-1), "CENTER"),
    ]))
    story.append(type_table)
    story.append(Spacer(1, 0.2*cm))

    # ── 待跟進 ───────────────────────────────────────────────────────────────
    story.append(Paragraph("▍ 下週待跟進", styles["section"]))
    df_t["下次跟進日期"] = pd.to_datetime(df_t["下次跟進日期"], errors="coerce")
    next_week = today + timedelta(days=7)
    due = df_t[
        df_t["下次跟進日期"].notna() &
        (df_t["下次跟進日期"] <= pd.Timestamp(next_week)) &
        (~df_t["狀態"].isin(["合作中","暫不合適"]))
    ].sort_values("下次跟進日期")

    if due.empty:
        story.append(Paragraph("下週無待跟進項目", styles["body"]))
    else:
        due_data = [[
            Paragraph("跟進日期", styles["tag"]),
            Paragraph("單位名稱", styles["tag"]),
            Paragraph("類型", styles["tag"]),
            Paragraph("目前狀態", styles["tag"]),
            Paragraph("備註", styles["tag"]),
        ]]
        for _, r in due.iterrows():
            days_left = (r["下次跟進日期"].date() - today.date()).days
            date_str  = r["下次跟進日期"].strftime("%m/%d")
            if days_left < 0:
                date_str = f"⚠ 逾期{abs(days_left)}天"
                ds = styles["bad"]
            elif days_left == 0:
                date_str = "今天"
                ds = styles["warn"]
            else:
                ds = styles["body"]
            due_data.append([
                Paragraph(date_str, ds),
                Paragraph(str(r["單位名稱"])[:18], styles["body"]),
                Paragraph(str(r["單位類型"]), styles["body"]),
                Paragraph(str(r["狀態"]), styles["body"]),
                Paragraph(str(r["備註"])[:25] if pd.notna(r["備註"]) else "", styles["body"]),
            ])
        cws3 = [(W-2*margin)*x for x in [0.16,0.26,0.18,0.16,0.24]]
        due_table = Table(due_data, colWidths=cws3)
        due_table.setStyle(TableStyle([
            ("BACKGROUND",    (0,0),(-1,0),  TEAL),
            ("TEXTCOLOR",     (0,0),(-1,0),  WHITE),
            ("ROWBACKGROUNDS",(0,1),(-1,-1), [LIGHT, colors.HexColor("#E8EFF8")]),
            ("LEFTPADDING",   (0,0),(-1,-1), 8),
            ("TOPPADDING",    (0,0),(-1,-1), 5),
            ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ]))
        story.append(due_table)

    # ── Footer ───────────────────────────────────────────────────────────────
    story.append(Spacer(1, 0.3*cm))
    story.append(HRFlowable(width=W-2*margin, color=TEAL, thickness=1))
    story.append(Paragraph(
        f"蒲朝棟 Tim Pu | CDA國際職涯諮詢師 | 104職涯引導師 | shoppy09@gmail.com | "
        f"報告生成時間：{today.strftime('%Y-%m-%d %H:%M')}",
        styles["footer"]
    ))

    doc.build(story)
    print(f"✓ 週報已生成：reports/{fname}")

    if open_after:
        import os
        os.startfile(str(out))
    return out


def print_weekly_summary():
    """在終端機顯示本週摘要"""
    df_t, df_l = load_data()
    today    = datetime.now()
    week_ago = today - timedelta(days=7)
    df_l["日期"] = pd.to_datetime(df_l["日期"], errors="coerce")
    this_week = df_l[df_l["日期"] >= week_ago]

    print("\n  ── 本週摘要 ──")
    print(f"  本週聯絡：{len(this_week)} 筆")
    print(f"  已發信：  {(df_t['狀態']=='已發信').sum()} 筆")
    print(f"  已回覆：  {(df_t['狀態']=='已回覆').sum()} 筆")
    print(f"  合作中：  {(df_t['狀態']=='合作中').sum()} 筆")
    print(f"  待聯絡：  {(df_t['狀態']=='未聯絡').sum()} 筆")


if __name__ == "__main__":
    generate_report(open_after=True)
