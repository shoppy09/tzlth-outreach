#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
合作提案書 PDF 生成模組
蒲朝棟 Tim Pu
"""

from pathlib import Path
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ── 字型設定（Windows / Linux 自動偵測）─────────────────────────────────────
_FONT_CANDIDATES = [
    # Windows
    ("C:/Windows/Fonts/msjh.ttc",  "C:/Windows/Fonts/msjhbd.ttc"),
    # Linux (Render / Ubuntu) - Noto CJK
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
     "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"),
    ("/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
     "/usr/share/fonts/noto-cjk/NotoSansCJK-Bold.ttc"),
    ("/usr/share/fonts/truetype/noto/NotoSansCJKtc-Regular.otf",
     "/usr/share/fonts/truetype/noto/NotoSansCJKtc-Bold.otf"),
    # Linux fallback - WQY
    ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
     "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"),
]

FONT_PATH, FONT_BOLD_PATH = _FONT_CANDIDATES[0]  # 預設 Windows

def register_fonts():
    import glob
    global FONT_PATH, FONT_BOLD_PATH
    all_candidates = list(_FONT_CANDIDATES)
    for regular, bold in all_candidates:
        if Path(regular).exists():
            FONT_PATH, FONT_BOLD_PATH = regular, bold
            bold_path = bold if Path(bold).exists() else regular
            try:
                pdfmetrics.registerFont(TTFont("MSJhei",   regular))
                pdfmetrics.registerFont(TTFont("MSJheiBd", bold_path))
                from reportlab.lib.fonts import addMapping
                addMapping("MSJhei",   0, 0, "MSJhei")
                addMapping("MSJhei",   1, 0, "MSJheiBd")
                addMapping("MSJhei",   0, 1, "MSJhei")
                addMapping("MSJhei",   1, 1, "MSJheiBd")
                addMapping("MSJheiBd", 0, 0, "MSJheiBd")
                addMapping("MSJheiBd", 1, 0, "MSJheiBd")
                addMapping("MSJheiBd", 0, 1, "MSJheiBd")
                addMapping("MSJheiBd", 1, 1, "MSJheiBd")
                return
            except Exception:
                continue
    # glob fallback: search entire /usr/share/fonts for any CJK font
    for pattern in ["/usr/share/fonts/**/*CJK*Regular*.ttc",
                    "/usr/share/fonts/**/*CJK*Regular*.otf",
                    "/usr/share/fonts/**/*[Nn]oto*[Cc][Jj][Kk]*.ttc"]:
        matches = glob.glob(pattern, recursive=True)
        if matches:
            try:
                from reportlab.lib.fonts import addMapping
                FONT_PATH = FONT_BOLD_PATH = matches[0]
                pdfmetrics.registerFont(TTFont("MSJhei",   matches[0]))
                pdfmetrics.registerFont(TTFont("MSJheiBd", matches[0]))
                addMapping("MSJhei",   0, 0, "MSJhei")
                addMapping("MSJhei",   1, 0, "MSJheiBd")
                addMapping("MSJhei",   0, 1, "MSJhei")
                addMapping("MSJhei",   1, 1, "MSJheiBd")
                addMapping("MSJheiBd", 0, 0, "MSJheiBd")
                addMapping("MSJheiBd", 1, 0, "MSJheiBd")
                addMapping("MSJheiBd", 0, 1, "MSJheiBd")
                addMapping("MSJheiBd", 1, 1, "MSJheiBd")
                return
            except Exception:
                continue

# ── 顏色 ──────────────────────────────────────────────────────────────────────
NAVY   = colors.HexColor("#1B3A6B")
TEAL   = colors.HexColor("#2A7D8C")
LIGHT  = colors.HexColor("#F0F5FA")
GREY   = colors.HexColor("#555555")
WHITE  = colors.white
BLACK  = colors.black

# ── 樣式 ──────────────────────────────────────────────────────────────────────
def make_styles():
    return {
        "name": ParagraphStyle(
            "name", fontName="MSJheiBd", fontSize=22,
            textColor=WHITE, spaceAfter=2, leading=28
        ),
        "subtitle": ParagraphStyle(
            "subtitle", fontName="MSJhei", fontSize=11,
            textColor=colors.HexColor("#CCE0FF"), spaceAfter=4, leading=16
        ),
        "section": ParagraphStyle(
            "section", fontName="MSJheiBd", fontSize=13,
            textColor=NAVY, spaceBefore=14, spaceAfter=6, leading=18
        ),
        "body": ParagraphStyle(
            "body", fontName="MSJhei", fontSize=10,
            textColor=GREY, spaceAfter=4, leading=16
        ),
        "bullet": ParagraphStyle(
            "bullet", fontName="MSJhei", fontSize=10,
            textColor=GREY, spaceAfter=3, leading=15, leftIndent=12
        ),
        "tag": ParagraphStyle(
            "tag", fontName="MSJheiBd", fontSize=9,
            textColor=TEAL, spaceAfter=2, leading=13
        ),
        "contact": ParagraphStyle(
            "contact", fontName="MSJhei", fontSize=10,
            textColor=WHITE, spaceAfter=3, leading=15
        ),
        "footer": ParagraphStyle(
            "footer", fontName="MSJhei", fontSize=8,
            textColor=colors.HexColor("#999999"), leading=12
        ),
    }

# ── 主生成函數 ────────────────────────────────────────────────────────────────
def generate_proposal(output_path: str | Path, target_name: str = "", target_type: str = "") -> Path:
    register_fonts()
    output_path = Path(output_path)
    styles = make_styles()
    W, H = A4
    margin = 1.8 * cm
    content_w = W - 2 * margin

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=margin, rightMargin=margin,
        topMargin=1.2 * cm, bottomMargin=1.5 * cm,
    )

    story = []

    # ── 頁首 Banner ─────────────────────────────────────────────────────────
    banner_data = [[
        Paragraph("蒲朝棟  Tim Pu", styles["name"]),
        ""
    ]]
    subtitle_data = [[
        Paragraph("職涯顧問｜CDA 國際職涯諮詢師｜104 職涯引導師", styles["subtitle"]),
        ""
    ]]
    banner_table = Table(
        [banner_data[0], subtitle_data[0]],
        colWidths=[content_w * 0.75, content_w * 0.25],
        rowHeights=[32, 22],
    )
    banner_table.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), NAVY),
        ("LEFTPADDING",  (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
        ("TOPPADDING",   (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 8),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(banner_table)
    story.append(Spacer(1, 0.3 * cm))

    # ── 標題列 ───────────────────────────────────────────────────────────────
    title_text = f"合作提案書"
    if target_name:
        title_text += f"  ·  {target_name}"
    title_para = Paragraph(title_text, ParagraphStyle(
        "title", fontName="MSJheiBd", fontSize=16,
        textColor=NAVY, spaceAfter=2, leading=22
    ))
    date_para = Paragraph(
        f"提案日期：{datetime.now().strftime('%Y 年 %m 月 %d 日')}",
        styles["footer"]
    )
    story.extend([title_para, date_para, Spacer(1, 0.2 * cm)])
    story.append(HRFlowable(width=content_w, color=TEAL, thickness=2, spaceAfter=10))

    # ── 關於我 ───────────────────────────────────────────────────────────────
    story.append(Paragraph("▍ 關於我", styles["section"]))

    about_rows = [
        ["姓名", "蒲朝棟  Tim Pu"],
        ["專業認證", "CDA 國際職涯諮詢師　104 職涯引導師"],
        ["工作經歷", "職業軍人 7 年 → 後勤主管 → 管理部主管（共 20+ 年）"],
        ["目前身份", "在職斜槓職涯顧問（現仍任職，非全職顧問）"],
        ["服務地點", "桃園（線上 / 現場皆可）"],
        ["網站", "https://tzlth-website.vercel.app/"],
    ]
    about_table = Table(
        [[Paragraph(r[0], styles["tag"]), Paragraph(r[1], styles["body"])]
         for r in about_rows],
        colWidths=[content_w * 0.22, content_w * 0.78],
    )
    about_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), LIGHT),
        ("BACKGROUND",    (0, 0), (0, -1), colors.HexColor("#DDE8F5")),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 10),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("ROWBACKGROUNDS",(0, 0), (-1, -1),
         [LIGHT, colors.HexColor("#E8EFF8")]),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(about_table)
    story.append(Spacer(1, 0.2 * cm))

    # ── 核心理念 ─────────────────────────────────────────────────────────────
    story.append(Paragraph("▍ 核心理念", styles["section"]))
    quote_table = Table(
        [[Paragraph(
            "「能力早就在那裡，差的是說清楚的語言。」",
            ParagraphStyle("quote", fontName="MSJheiBd", fontSize=12,
                           textColor=TEAL, leading=20)
        )]],
        colWidths=[content_w],
    )
    quote_table.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, -1), colors.HexColor("#EAF4F6")),
        ("LEFTPADDING",  (0, 0), (-1, -1), 16),
        ("TOPPADDING",   (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 10),
        ("LINEAFTER",    (0, 0), (0, -1), 4, TEAL),
    ]))
    story.append(quote_table)
    story.append(Paragraph(
        "許多人卡住不是因為能力不足，而是不知道如何把已有的能力轉譯成市場能理解的語言。"
        "我的工作是幫助每一位學員找到屬於自己的職涯語言，讓他們的價值被正確看見。",
        styles["body"]
    ))

    # ── 可提供的服務 ─────────────────────────────────────────────────────────
    story.append(Paragraph("▍ 可提供的服務", styles["section"]))

    services = [
        ("職涯轉型工作坊",   "1.5–3 小時｜10–30 人｜NT$8,000–20,000"),
        ("履歷 × 面試工作坊", "2–3 小時｜15–25 人｜NT$12,000–20,000"),
        ("1 對 1 職涯諮詢",  "60–90 分鐘｜NT$1,500–2,000 / 次"),
        ("職涯陪跑方案",     "每月 4 次｜NT$6,000 / 月"),
        ("企業內訓 / 員工職涯發展", "客製化報價"),
        ("演講 / 業師計畫",  "依場次與學校規範洽談"),
    ]
    svc_data = [
        [Paragraph(f"◆ {s[0]}", ParagraphStyle(
            "svc_title", fontName="MSJheiBd", fontSize=10,
            textColor=NAVY, leading=15)),
         Paragraph(s[1], styles["body"])]
        for s in services
    ]
    svc_table = Table(svc_data, colWidths=[content_w * 0.38, content_w * 0.62])
    svc_table.setStyle(TableStyle([
        ("ROWBACKGROUNDS",  (0, 0), (-1, -1),
         [LIGHT, colors.HexColor("#E8EFF8")]),
        ("LEFTPADDING",     (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",    (0, 0), (-1, -1), 10),
        ("TOPPADDING",      (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING",   (0, 0), (-1, -1), 6),
        ("VALIGN",          (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(svc_table)
    story.append(Spacer(1, 0.2 * cm))

    # ── 課程大綱（依對象類型客製） ─────────────────────────────────────────────
    CURRICULA = {
        "大專院校": [
            ("職涯探索講座（90 分鐘）",
             "單元一：你以為你不知道，但你早就知道了（職能盤點）\n"
             "單元二：履歷不是人生回顧，是行銷工具\n"
             "單元三：面試就是說故事，你的故事說得動人嗎？\n"
             "形式：演講 + Q&A｜適合大班（50人+）"),
            ("履歷×面試工作坊（3 小時）",
             "第一段：你的強項在哪裡？職能三角形盤點法\n"
             "第二段：把經驗寫成成果——STAR法則實作\n"
             "第三段：模擬面試 + 即時回饋\n"
             "形式：互動工作坊｜小班（15–25人）"),
        ],
        "社區大學": [
            ("職涯第二春——35歲後的轉職策略（6堂）",
             "第1堂：職涯健檢——找出卡住的真正原因\n"
             "第2堂：能力盤點——你比你想的更有價值\n"
             "第3堂：市場定位——你的目標客群是誰\n"
             "第4堂：履歷重寫——讓經歷說對的語言\n"
             "第5堂：面試心法——從緊張到從容\n"
             "第6堂：行動計畫——下一步怎麼走"),
            ("斜槓起步（4堂）",
             "第1堂：找到你的斜槓方向\n"
             "第2堂：定價與市場測試\n"
             "第3堂：個人品牌建立\n"
             "第4堂：從第一個客戶開始"),
        ],
        "退輔會/榮服處": [
            ("退役軍人職涯轉型工作坊（3 小時）",
             "第一段：把軍中能力翻譯成企業語言\n"
             "第二段：退役軍人的履歷寫法\n"
             "第三段：面試實戰——讓HR看見你的價值\n"
             "形式：工作坊 + 個人練習｜10–20人"),
        ],
        "企業HR": [
            ("員工職涯探索工作坊（3 小時）",
             "第一段：為什麼我在這裡？——職涯意義探索\n"
             "第二段：我能做什麼？——職能盤點與差距分析\n"
             "第三段：下一步怎麼走？——個人發展計畫制定\n"
             "形式：工作坊｜10–20人｜可含個別諮詢"),
        ],
    }

    if target_type in CURRICULA:
        story.append(Paragraph("▍ 課程大綱", styles["section"]))
        for course_name, outline in CURRICULA[target_type]:
            story.append(Paragraph(f"◆ {course_name}", ParagraphStyle(
                "course", fontName="MSJheiBd", fontSize=10,
                textColor=NAVY, spaceAfter=4, leading=15)))
            outline_table = Table([[Paragraph(outline, styles["body"])]],
                                  colWidths=[content_w])
            outline_table.setStyle(TableStyle([
                ("BACKGROUND",    (0,0),(-1,-1), LIGHT),
                ("LEFTPADDING",   (0,0),(-1,-1), 12),
                ("RIGHTPADDING",  (0,0),(-1,-1), 12),
                ("TOPPADDING",    (0,0),(-1,-1), 8),
                ("BOTTOMPADDING", (0,0),(-1,-1), 8),
                ("LINEAFTER",     (0,0),(0,-1),  2, TEAL),
            ]))
            story.append(outline_table)
            story.append(Spacer(1, 0.15*cm))

    # ── 合作流程 ─────────────────────────────────────────────────────────────
    story.append(Paragraph("▍ 合作流程", styles["section"]))
    steps = [
        ("Step 1", "初步洽談", "電話 / Email / LINE 說明需求與期望"),
        ("Step 2", "提案確認", "提供課程大綱、師資資料、費用說明"),
        ("Step 3", "簽訂合約", "明確時間、場地、費用、開立發票"),
        ("Step 4", "課前準備", "客製化課程內容，配合受眾背景調整"),
        ("Step 5", "執行課程", "現場 / 線上授課，提供教材"),
        ("Step 6", "課後追蹤", "滿意度回饋，持續合作評估"),
    ]
    step_data = [[
        Paragraph(s[0], ParagraphStyle(
            "step_num", fontName="MSJheiBd", fontSize=9,
            textColor=WHITE, leading=13, alignment=1)),
        Paragraph(f"<b>{s[1]}</b><br/>{s[2]}", ParagraphStyle(
            "step_body", fontName="MSJhei", fontSize=9,
            textColor=GREY, leading=14))
    ] for s in steps]
    step_table = Table(step_data, colWidths=[content_w * 0.15, content_w * 0.85])
    step_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (0, -1), TEAL),
        ("BACKGROUND",    (1, 0), (1, -1), LIGHT),
        ("ROWBACKGROUNDS",(1, 0), (1, -1),
         [LIGHT, colors.HexColor("#E8EFF8")]),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN",         (0, 0), (0, -1), "CENTER"),
    ]))
    story.append(step_table)

    # ── 聯絡資訊 ─────────────────────────────────────────────────────────────
    story.append(Spacer(1, 0.3 * cm))
    contact_data = [[
        Paragraph("📧  shoppy09@gmail.com", styles["contact"]),
        Paragraph("🌐  tzlth-website.vercel.app", styles["contact"]),
        Paragraph("💬  LINE：lin.ee/IOX6V66", styles["contact"]),
        Paragraph("📍  桃園（線上 / 面談皆可）", styles["contact"]),
    ]]
    contact_table = Table(
        contact_data,
        colWidths=[content_w * 0.28, content_w * 0.28,
                   content_w * 0.26, content_w * 0.18],
    )
    contact_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), NAVY),
        ("LEFTPADDING",   (0, 0), (-1, -1), 10),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(contact_table)

    doc.build(story)
    return output_path


if __name__ == "__main__":
    out = Path("outputs/test_proposal.pdf")
    out.parent.mkdir(exist_ok=True)
    generate_proposal(out, target_name="測試單位", target_type="大專院校")
    print(f"PDF 生成完成：{out}")
