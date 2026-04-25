#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合作提案書生成工具 — 非互動式 Claude Code 整合版

使用方式：
  # 僅生成 PDF（Tim 確認後再發送）
  python scripts/generate_proposal.py --name "元智大學職涯發展中心" --type "大專院校"

  # 生成並立即發送（需 config.env 已設定 SENDER_EMAIL / SENDER_PASSWORD）
  python scripts/generate_proposal.py --name "元智大學職涯發展中心" --type "大專院校" \
      --email "career@yzu.edu.tw" --send

有效的 --type 選項：
  大專院校 / 社區大學 / 退輔會/榮服處 / 企業HR / 政府勞動局處 / 人資協會

與 main.py 的關係：
  此腳本為非互動式 wrapper，不依賴 main.py 的 REPL 選單。
  直接呼叫 pdf_gen.generate_proposal() 生成 PDF，
  並複用 main.py 的 load_cfg() + send_email() 發送含附件 Email。
"""
import sys
import argparse
from pathlib import Path
from datetime import datetime

# 讓 Python 找得到上層的 pdf_gen.py 與 main.py
BASE_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(BASE_DIR))

import pdf_gen
from main import load_cfg, send_email  # send_email 已支援 MIMEMultipart 附件

OUTPUTS_DIR = BASE_DIR / "outputs"
OUTPUTS_DIR.mkdir(exist_ok=True)

VALID_TYPES = [
    "大專院校",
    "社區大學",
    "退輔會/榮服處",
    "企業HR",
    "政府勞動局處",
    "人資協會",
]


def generate_pdf(name: str, unit_type: str) -> Path:
    """生成合作提案書 PDF，回傳檔案路徑。"""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = name.replace("/", "_").replace(" ", "_").replace("　", "_")
    output_path = OUTPUTS_DIR / f"{ts}_{safe_name}_合作提案書.pdf"
    result = pdf_gen.generate_proposal(output_path, target_name=name, target_type=unit_type)
    print(f"✅ PDF 已生成：{result}")
    return result


def compose_email_body(unit_name: str, unit_type: str) -> tuple[str, str]:
    """依單位類型組合 Email 主旨與內文。"""
    if unit_type in ("大專院校",):
        subject = f"【職涯講座合作提案】蒲朝棟顧問 × {unit_name}"
        body = (
            f"您好，\n\n"
            f"敝人蒲朝棟，為 CDA 國際認證職涯諮詢師暨 104 職涯引導師，"
            f"專注協助 3–10 年職場工作者解決職涯轉型與求職問題。\n\n"
            f"誠摯邀請 {unit_name} 職涯發展中心考慮合作舉辦職涯探索或求職實戰講座，"
            f"詳細提案書請見附件，期待有機會進一步交流！\n\n"
            f"蒲朝棟 Tim Pu\n"
            f"📧 shoppy09@gmail.com\n"
            f"💬 LINE：lin.ee/IOX6V66\n"
            f"🌐 www.careerssl.com"
        )
    elif unit_type in ("社區大學",):
        subject = f"【課程合作提案】蒲朝棟顧問 × {unit_name}"
        body = (
            f"您好，\n\n"
            f"敝人蒲朝棟，CDA 認證職涯顧問，長期協助職場工作者規劃職涯方向。\n\n"
            f"誠摯建議 {unit_name} 考慮開設「職涯停看聽」系列課程，"
            f"詳細提案書請見附件，期待有機會進一步討論！\n\n"
            f"蒲朝棟 Tim Pu\n"
            f"📧 shoppy09@gmail.com\n"
            f"💬 LINE：lin.ee/IOX6V66"
        )
    elif unit_type in ("退輔會/榮服處",):
        subject = f"【退役軍人職涯輔導合作提案】蒲朝棟顧問 × {unit_name}"
        body = (
            f"您好，\n\n"
            f"敝人蒲朝棟，CDA 國際認證職涯諮詢師，"
            f"對退役軍人轉職輔導有豐富經驗。\n\n"
            f"附上合作提案書供參考，誠摯期待與貴單位攜手提升退役弟兄的職涯競爭力！\n\n"
            f"蒲朝棟 Tim Pu\n"
            f"📧 shoppy09@gmail.com"
        )
    elif unit_type in ("企業HR",):
        subject = f"【員工職涯賦能講座提案】蒲朝棟顧問 × {unit_name}"
        body = (
            f"您好，\n\n"
            f"敝人蒲朝棟，CDA 認證職涯顧問，"
            f"專長協助企業員工在職涯發展與轉型期做出清晰決策。\n\n"
            f"附上合作提案書供參考，歡迎進一步洽談！\n\n"
            f"蒲朝棟 Tim Pu\n"
            f"📧 shoppy09@gmail.com"
        )
    else:
        # 政府勞動局處 / 人資協會 / 其他
        subject = f"【合作提案】蒲朝棟職涯顧問 × {unit_name}"
        body = (
            f"您好，\n\n"
            f"敝人蒲朝棟，CDA 認證職涯諮詢師，附上合作提案書供參考。\n\n"
            f"期待有機會進一步交流！\n\n"
            f"蒲朝棟 Tim Pu\n"
            f"📧 shoppy09@gmail.com\n"
            f"💬 LINE：lin.ee/IOX6V66"
        )
    return subject, body


def main():
    parser = argparse.ArgumentParser(
        description="合作提案書 PDF 生成 + Email 發送工具（Claude Code 非互動式整合版）"
    )
    parser.add_argument(
        "--name", required=True,
        help="單位名稱（例：元智大學職涯發展中心）"
    )
    parser.add_argument(
        "--type", required=True, choices=VALID_TYPES, dest="unit_type",
        help="單位類型（大專院校 / 社區大學 / 退輔會/榮服處 / 企業HR / 政府勞動局處 / 人資協會）"
    )
    parser.add_argument(
        "--email", default="",
        help="收件人 Email（不填則只生成 PDF，不發送）"
    )
    parser.add_argument(
        "--send", action="store_true",
        help="生成後立即發送 Email（需同時指定 --email）"
    )
    args = parser.parse_args()

    # Step 1：生成 PDF
    pdf_path = generate_pdf(args.name, args.unit_type)

    # Step 2：若指定 --send，則自動組信並發送
    if args.send:
        if not args.email:
            print("❌ 使用 --send 時必須同時指定 --email")
            sys.exit(1)
        subject, body = compose_email_body(args.name, args.unit_type)
        success = send_email(
            to=args.email,
            subj=subject,
            body=body,
            pdf_path=str(pdf_path),
        )
        sys.exit(0 if success else 1)
    else:
        # 僅生成，等 Tim 確認
        print()
        print("📋 Tim 確認後，執行以下指令發送：")
        print(
            f"  python scripts/generate_proposal.py"
            f' --name "{args.name}"'
            f' --type "{args.unit_type}"'
            f' --email "[收件 Email]"'
            f" --send"
        )
        print()
        print(f"📄 PDF 已存至：{pdf_path}")


if __name__ == "__main__":
    main()
