#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
職涯顧問合作拓展系統 v2.0
蒲朝棟 Tim Pu | CDA 國際職涯諮詢師 | 104 職涯引導師
"""
import os, sys, smtplib
from pathlib import Path
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8","utf8"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

import pandas as pd

BASE_DIR      = Path(__file__).parent
DATA_FILE     = BASE_DIR / "data" / "targets.xlsx"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUTS_DIR   = BASE_DIR / "outputs"
CONFIG_FILE   = BASE_DIR / "config.env"

# ── 顏色 ─────────────────────────────────────────────────────────────────────
R="\033[91m"; G="\033[92m"; Y="\033[93m"
B="\033[94m"; C="\033[96m"; BOLD="\033[1m"; Z="\033[0m"

TARGET_SHEET = "目標單位"
LOG_SHEET    = "聯絡記錄"

STATUSES      = ["未聯絡","已發信","已回覆","已安排會議","合作中","暫不合適"]
TYPES         = ["退輔會/榮服處","政府勞動局處","大專院校","社區大學","企業HR","人資協會"]
CONTACT_MODES = ["Email","電話先行","線上表單","LINE"]
PRIORITIES    = ["高","中","低"]
ROUNDS        = ["第一輪","第二輪"]

TYPE_TMPL = {
    ("退輔會/榮服處","第一輪"): "退輔會_第一輪.txt",
    ("退輔會/榮服處","第二輪"): "退輔會_第二輪.txt",
    ("政府勞動局處","第一輪"): "政府勞動局_第一輪.txt",
    ("政府勞動局處","第二輪"): "政府勞動局_第二輪.txt",
    ("大專院校","第一輪"):     "大專院校_第一輪.txt",
    ("大專院校","第二輪"):     "大專院校_第二輪.txt",
    ("社區大學","第一輪"):     "社區大學_第一輪.txt",
    ("社區大學","第二輪"):     "社區大學_第二輪.txt",
    ("企業HR","第一輪"):       "企業HR_第一輪.txt",
    ("企業HR","第二輪"):       "企業HR_第二輪.txt",
    ("人資協會","第一輪"):     "人資協會_第一輪.txt",
    ("人資協會","第二輪"):     "人資協會_第二輪.txt",
}

# ── 預載資料 ──────────────────────────────────────────────────────────────────
def _t(id,name,typ,phone,region,priority,mode,note,email=""):
    return {"ID":id,"單位名稱":name,"單位類型":typ,"聯絡人":"","職稱":"",
            "Email":email,"電話":phone,"地區":region,"狀態":"未聯絡",
            "優先序":priority,"聯絡方式":mode,"輪次":"第一輪",
            "發信日期":None,"最後跟進日期":None,"下次跟進日期":None,"備註":note}

SAMPLE = [
    # 退輔會
    _t(1,"國軍退除役官兵輔導委員會","退輔會/榮服處","02-2595-1141","台北","高","電話先行","全國退輔會總部，洽詢職訓合作"),
    _t(2,"桃園市政府榮民服務處","退輔會/榮服處","03-336-2280","桃園","高","電話先行","在地優先，退役軍人轉職輔導"),
    _t(3,"新北市政府榮民服務處","退輔會/榮服處","02-2960-3456","新北","中","電話先行",""),
    # 政府勞動
    _t(4,"勞動力發展署桃竹苗分署","政府勞動局處","03-462-1681","桃園","高","Email","有職訓課程合作預算",""),
    _t(5,"桃園市政府勞動局就業服務科","政府勞動局處","03-332-5632","桃園","高","Email","在地就業服務","labor@tycg.gov.tw"),
    _t(6,"桃園市就業服務站","政府勞動局處","03-337-7333","桃園","中","電話先行",""),
    # 大學（桃園）
    _t(7,"元智大學職涯發展中心","大專院校","03-463-8800","桃園","高","Email","桃園在地，可邀業師"),
    _t(8,"中原大學職涯發展中心","大專院校","03-265-7741","桃園","高","Email","理工科學生職涯需求強"),
    _t(9,"長庚大學職涯發展中心","大專院校","03-211-8800","桃園","高","Email",""),
    _t(10,"銘傳大學職涯發展中心","大專院校","03-350-7001","桃園","高","Email",""),
    _t(11,"國立中央大學職涯發展中心","大專院校","03-422-7151","桃園","高","Email","國立大學，聲望佳"),
    _t(12,"健行科技大學職涯發展中心","大專院校","03-428-6555","桃園","中","Email",""),
    _t(13,"龍華科技大學職涯發展中心","大專院校","03-411-7580","桃園","中","Email",""),
    _t(14,"開南大學職涯發展中心","大專院校","03-341-2500","桃園","中","Email",""),
    _t(15,"萬能科技大學職涯發展中心","大專院校","03-451-5811","桃園","低","Email",""),
    # 大學（台北）
    _t(16,"國立台灣科技大學職涯中心","大專院校","02-2733-3141","台北","中","Email","理工強校，職涯需求大"),
    _t(17,"國立台北科技大學職涯中心","大專院校","02-2771-2171","台北","中","Email",""),
    _t(18,"輔仁大學職涯發展中心","大專院校","02-2905-2693","新北","中","Email",""),
    _t(19,"淡江大學職涯與就業輔導組","大專院校","02-2621-5656","新北","中","Email",""),
    _t(20,"文化大學職涯中心","大專院校","02-2861-0511","台北","中","Email",""),
    _t(21,"東吳大學職涯發展中心","大專院校","02-2311-1531","台北","中","Email",""),
    _t(22,"世新大學職涯發展中心","大專院校","02-2236-8225","台北","中","Email",""),
    _t(23,"國立台北大學職涯發展中心","大專院校","02-8674-1111","新北","中","Email",""),
    _t(24,"國立政治大學職涯中心","大專院校","02-2939-3091","台北","低","Email","競爭者多，但聲望高"),
    # 大學（新竹）
    _t(25,"國立清華大學職涯中心","大專院校","03-571-5131","新竹","中","Email",""),
    _t(26,"國立陽明交通大學職涯中心","大專院校","03-571-2121","新竹","中","Email",""),
    _t(27,"中華大學職涯發展中心","大專院校","03-518-6404","新竹","低","Email",""),
    # 社區大學
    _t(28,"桃園市立中壢社區大學","社區大學","03-425-3106","桃園","高","Email","成人學員，職涯第二春需求高","jhongli@tycg.gov.tw"),
    _t(29,"桃園市立桃園社區大學","社區大學","03-356-1688","桃園","高","Email",""),
    _t(30,"桃園市立蘆竹社區大學","社區大學","03-211-3877","桃園","中","Email",""),
    _t(31,"新北市新莊社區大學","社區大學","02-2908-5225","新北","低","Email",""),
    # 人資協會
    _t(32,"台灣人力資源管理協會","人資協會","02-2708-2268","台北","高","Email","投稿或申請演講"),
    _t(33,"中華人力資源管理協會","人資協會","02-2562-3737","台北","高","Email",""),
]

TARGET_COLS = ["ID","單位名稱","單位類型","聯絡人","職稱","Email","電話",
               "地區","狀態","優先序","聯絡方式","輪次",
               "發信日期","最後跟進日期","下次跟進日期","備註"]
LOG_COLS    = ["日期","單位ID","單位名稱","聯絡方式","內容摘要","結果"]

# ── 資料管理 ──────────────────────────────────────────────────────────────────
BACKUP_DIR = BASE_DIR / "data" / "backup"

def ensure_dirs():
    (BASE_DIR/"data").mkdir(exist_ok=True)
    BACKUP_DIR.mkdir(exist_ok=True)
    TEMPLATES_DIR.mkdir(exist_ok=True)
    OUTPUTS_DIR.mkdir(exist_ok=True)

def init_data():
    if DATA_FILE.exists():
        # 若舊資料欠新欄位，補上
        df = pd.read_excel(DATA_FILE, sheet_name=TARGET_SHEET)
        changed = False
        for col, default in [("優先序","中"),("聯絡方式","Email"),("輪次","第一輪")]:
            if col not in df.columns:
                df[col] = default
                changed = True
        if changed:
            df_l = pd.read_excel(DATA_FILE, sheet_name=LOG_SHEET)
            _save(df[TARGET_COLS], df_l)
            print(f"{G}✓ 資料庫欄位已更新{Z}")
        return
    df_t = pd.DataFrame(SAMPLE, columns=TARGET_COLS)
    df_l = pd.DataFrame(columns=LOG_COLS)
    _save(df_t, df_l)
    print(f"{G}✓ 資料庫初始化完成，預載 {len(SAMPLE)} 筆目標單位{Z}")

def load_t() -> pd.DataFrame:
    return pd.read_excel(DATA_FILE, sheet_name=TARGET_SHEET, dtype={"ID":int})

def load_l() -> pd.DataFrame:
    return pd.read_excel(DATA_FILE, sheet_name=LOG_SHEET)

def _save(df_t, df_l):
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
        df_t.to_excel(w, sheet_name=TARGET_SHEET, index=False)
        df_l.to_excel(w, sheet_name=LOG_SHEET, index=False)
    # 自動同步到 GitHub（Render 雲端持久化）
    try:
        from github_sync import sync_targets
        import threading
        threading.Thread(target=sync_targets, args=(DATA_FILE,), daemon=True).start()
    except Exception:
        pass
    # 自動備份：每天保留一份，以日期命名
    try:
        import shutil
        stamp = datetime.now().strftime("%Y%m%d")
        backup_path = BACKUP_DIR / f"targets_{stamp}.xlsx"
        if not backup_path.exists():
            shutil.copy2(DATA_FILE, backup_path)
            # 只保留最近 30 份
            old = sorted(BACKUP_DIR.glob("targets_*.xlsx"))[:-30]
            for f in old:
                f.unlink()
    except Exception:
        pass

def next_id(df): return int(df["ID"].max())+1 if not df.empty else 1

# ── UI ────────────────────────────────────────────────────────────────────────
def clr(): os.system("cls" if os.name=="nt" else "clear")
def hr(c="─",w=62): print(C+c*w+Z)

def header(t):
    hr("═"); print(f"{BOLD}{C}  {t}{Z}"); hr("═")

def sfmt(s):
    m={"未聯絡":"","已發信":Y,"已回覆":B,"已安排會議":C,"合作中":G,"暫不合適":R}
    return f"{m.get(s,'')}{s}{Z}"

def pfmt(p):
    return f"{R}{p}{Z}" if p=="高" else (f"{Y}{p}{Z}" if p=="中" else p)

def pick(opts, prompt="選擇編號"):
    for i,o in enumerate(opts,1): print(f"  {i}. {o}")
    while True:
        try:
            idx=int(input(f"{prompt}: ").strip())-1
            if 0<=idx<len(opts): return idx
        except ValueError: pass
        print(f"{R}請輸入有效編號{Z}")

# ── 查看 ──────────────────────────────────────────────────────────────────────
def view_targets(df=None, fs=None, ft=None, fr=None, fp=None):
    if df is None: df=load_t()
    if fs: df=df[df["狀態"]==fs]
    if ft: df=df[df["單位類型"]==ft]
    if fr: df=df[df["地區"].str.contains(fr,na=False)]
    if fp: df=df[df["優先序"]==fp]
    if df.empty: print(f"{Y}  沒有符合條件的記錄{Z}"); return
    print(f"\n{'ID':>3}  {'單位名稱':<20}  {'類型':<12}  {'地區':<5}  {'優先':4}  {'輪次':<5}  狀態")
    hr()
    for _,r in df.iterrows():
        print(f"{int(r['ID']):>3}  {str(r['單位名稱']):<20}  {str(r['單位類型']):<12}  "
              f"{str(r['地區']):<5}  {pfmt(str(r.get('優先序','中'))):8}  "
              f"{str(r.get('輪次','第一輪')):<5}  {sfmt(str(r['狀態']))}")
    print(f"\n  共 {len(df)} 筆")

# ── 新增 ──────────────────────────────────────────────────────────────────────
def add_target():
    df_t=load_t(); df_l=load_l()
    header("新增目標單位")
    nid=next_id(df_t)
    print("\n單位類型："); ti=pick(TYPES)
    name    = input("單位名稱: ").strip()
    contact = input("聯絡人 (可空白): ").strip()
    email   = input("Email (可空白): ").strip()
    phone   = input("電話 (可空白): ").strip()
    region  = input("地區: ").strip()
    print("優先序："); pi=pick(PRIORITIES)
    print("聯絡方式："); ci=pick(CONTACT_MODES)
    note    = input("備註: ").strip()
    row={"ID":nid,"單位名稱":name,"單位類型":TYPES[ti],"聯絡人":contact,"職稱":"",
         "Email":email,"電話":phone,"地區":region,"狀態":"未聯絡",
         "優先序":PRIORITIES[pi],"聯絡方式":CONTACT_MODES[ci],"輪次":"第一輪",
         "發信日期":None,"最後跟進日期":None,"下次跟進日期":None,"備註":note}
    df_t=pd.concat([df_t,pd.DataFrame([row])],ignore_index=True)
    _save(df_t,df_l)
    print(f"{G}✓ 已新增：{name}（ID: {nid}）{Z}")

# ── 模板 ──────────────────────────────────────────────────────────────────────
def load_tmpl(unit_type, round_):
    fname=TYPE_TMPL.get((unit_type, round_))
    if not fname: return None
    p=TEMPLATES_DIR/fname
    return p.read_text(encoding="utf-8") if p.exists() else None

def fill_tmpl(tmpl, row):
    import pandas as pd
    contact=str(row["聯絡人"]) if pd.notna(row["聯絡人"]) and str(row["聯絡人"]).strip() else ""
    salute=f"{contact}您好" if contact else "您好"
    return (tmpl
        .replace("{{單位名稱}}", str(row["單位名稱"]))
        .replace("{{聯絡人}}",   contact or "您")
        .replace("{{稱謂}}",     salute)
        .replace("{{地區}}",     str(row["地區"]) if pd.notna(row["地區"]) else "")
    )

# ── 單筆生成信件 ───────────────────────────────────────────────────────────────
def generate_one(target_id=None, silent=False):
    df_t=load_t(); df_l=load_l()
    if not silent: header("生成個人化信件")

    if target_id is None:
        if not silent:
            print("\n篩選「未聯絡」的單位：")
            view_targets(df_t, fs="未聯絡")
        target_id=int(input("\n輸入目標單位 ID: ").strip())

    rows=df_t[df_t["ID"]==target_id]
    if rows.empty: print(f"{R}找不到 ID {target_id}{Z}"); return None,None
    r=rows.iloc[0]

    round_=str(r.get("輪次","第一輪"))
    tmpl=load_tmpl(str(r["單位類型"]), round_)
    if not tmpl:
        print(f"{R}找不到模板：{r['單位類型']} {round_}{Z}"); return None,None

    body=fill_tmpl(tmpl, r)
    fname=f"{datetime.now().strftime('%Y%m%d')}_{r['單位名稱']}_{round_}.txt"
    out=OUTPUTS_DIR/fname
    out.write_text(body, encoding="utf-8")

    if not silent:
        print(f"\n{C}{'─'*60}{Z}")
        print(body)
        print(f"{C}{'─'*60}{Z}")
        print(f"\n{G}✓ 已儲存：outputs/{fname}{Z}")

    return body, r

def _mark_sent(df_t, df_l, target_id, r, note="發送合作邀請信"):
    today  = datetime.now().strftime("%Y-%m-%d")
    follow = (datetime.now()+timedelta(days=7)).strftime("%Y-%m-%d")
    for col in ["狀態", "發信日期", "下次跟進日期"]:
        if col in df_t.columns:
            df_t[col] = df_t[col].astype(object)
    df_t.loc[df_t["ID"]==target_id, ["狀態","發信日期","下次跟進日期"]] = ["已發信",today,follow]
    log={"日期":today,"單位ID":target_id,"單位名稱":r["單位名稱"],
         "聯絡方式":"Email","內容摘要":note,"結果":"待回覆"}
    return df_t, pd.concat([df_l,pd.DataFrame([log])],ignore_index=True)

def generate_email_interactive():
    body,r = generate_one()
    if body is None: return
    df_t=load_t(); df_l=load_l()
    target_id=int(df_t[df_t["單位名稱"]==r["單位名稱"]].iloc[0]["ID"])

    # 生成 PDF
    gen_pdf=input("\n是否同時生成合作提案書 PDF？(y/n): ").strip().lower()
    pdf_path=None
    if gen_pdf=="y":
        pdf_path=_gen_pdf(r)

    # 發送
    if pd.notna(r["Email"]) and str(r["Email"]).strip():
        send=input(f"\n偵測到 Email：{r['Email']}，是否發送？(y/n): ").strip().lower()
        if send=="y":
            subj=_subject(r)
            send_email(str(r["Email"]), subj, body, pdf_path)

    # 更新狀態
    upd=input("\n是否將狀態更新為「已發信」？(y/n): ").strip().lower()
    if upd=="y":
        df_t,df_l=_mark_sent(df_t,df_l,target_id,r)
        _save(df_t,df_l)
        print(f"{G}✓ 狀態已更新{Z}")

def _subject(r):
    round_=str(r.get("輪次","第一輪"))
    typ=str(r["單位類型"])
    if round_=="第二輪":
        return f"合作提案書附上｜{typ}職涯課程｜蒲朝棟"
    subjects={
        "退輔會/榮服處":"職涯轉型課程合作提案｜CDA 國際職涯諮詢師 蒲朝棟",
        "政府勞動局處":"職涯諮詢課程合作提案｜CDA 國際職涯諮詢師 蒲朝棟",
        "大專院校":    "職涯講師合作邀請｜CDA 國際職涯諮詢師 蒲朝棟",
        "社區大學":    "職涯課程合作提案｜CDA 國際職涯諮詢師 蒲朝棟",
        "企業HR":      "員工職涯發展課程提案｜CDA 國際職涯諮詢師 蒲朝棟",
        "人資協會":    "講師投稿合作邀請｜CDA 國際職涯諮詢師 蒲朝棟",
    }
    return subjects.get(typ,"合作提案｜CDA 國際職涯諮詢師 蒲朝棟")

def _gen_pdf(r):
    try:
        from pdf_gen import generate_proposal
        fname=f"{datetime.now().strftime('%Y%m%d')}_合作提案書_{r['單位名稱']}.pdf"
        path=OUTPUTS_DIR/fname
        generate_proposal(path, target_name=str(r["單位名稱"]), target_type=str(r["單位類型"]))
        print(f"{G}✓ PDF 已生成：outputs/{fname}{Z}")
        return path
    except Exception as e:
        print(f"{Y}⚠ PDF 生成失敗：{e}{Z}")
        return None

# ── 批次發信 ──────────────────────────────────────────────────────────────────
def batch_send():
    df_t=load_t(); df_l=load_l()
    header("批次發信")

    print("\n篩選條件（直接 Enter 跳過）：")
    fs=input(f"狀態 (預設:未聯絡): ").strip() or "未聯絡"
    ft=input(f"類型 [{'/'.join(TYPES)}]: ").strip() or None
    fr=input("地區 (如桃園): ").strip() or None
    fp=input("優先序 [高/中/低]: ").strip() or None

    targets=df_t.copy()
    if fs: targets=targets[targets["狀態"]==fs]
    if ft: targets=targets[targets["單位類型"]==ft]
    if fr: targets=targets[targets["地區"].str.contains(fr,na=False)]
    if fp: targets=targets[targets["優先序"]==fp]

    # 只保留有 Email 或可發信的
    has_email=targets[targets["Email"].notna() & (targets["Email"].astype(str).str.strip()!="")]
    no_email =targets[targets["Email"].isna() | (targets["Email"].astype(str).str.strip()=="")]

    print(f"\n{G}有 Email 可直接發送：{len(has_email)} 筆{Z}")
    print(f"{Y}缺 Email（僅生成文字檔）：{len(no_email)} 筆{Z}")

    if targets.empty:
        print(f"{R}沒有符合條件的單位{Z}"); return

    view_targets(targets)
    confirm=input(f"\n確認對以上 {len(targets)} 筆單位生成信件？(y/n): ").strip().lower()
    if confirm!="y": return

    gen_pdf_all=input("是否同時生成每份合作提案書 PDF？(y/n): ").strip().lower()
    send_all   =input("有 Email 的單位是否直接發送？(y/n): ").strip().lower()

    ok,fail=0,0
    for _,r in targets.iterrows():
        tid=int(r["ID"])
        round_=str(r.get("輪次","第一輪"))
        tmpl=load_tmpl(str(r["單位類型"]),round_)
        if not tmpl:
            print(f"  {Y}⚠ ID {tid} {r['單位名稱']}：找不到模板，跳過{Z}")
            fail+=1; continue

        body=fill_tmpl(tmpl,r)
        fname=f"{datetime.now().strftime('%Y%m%d')}_{r['單位名稱']}_{round_}.txt"
        (OUTPUTS_DIR/fname).write_text(body,encoding="utf-8")

        pdf_path=None
        if gen_pdf_all=="y":
            pdf_path=_gen_pdf(r)

        email_str=str(r["Email"]) if pd.notna(r["Email"]) else ""
        sent=False
        if send_all=="y" and email_str.strip():
            subj=_subject(r)
            sent=send_email(email_str,subj,body,pdf_path,quiet=True)

        df_t,df_l=_mark_sent(df_t,df_l,tid,r)
        status_mark=f"{G}✓ 已發送{Z}" if sent else f"{B}✓ 已生成{Z}"
        print(f"  {status_mark}  ID {tid}  {r['單位名稱']}")
        ok+=1

    _save(df_t,df_l)
    print(f"\n{G}完成！成功 {ok} 筆，失敗 {fail} 筆{Z}")
    print(f"信件檔案已存至 outputs/ 資料夾")

# ── 更新狀態 ──────────────────────────────────────────────────────────────────
def update_status():
    df_t=load_t(); df_l=load_l()
    header("更新聯絡狀態")
    view_targets(df_t)

    tid=int(input("\n輸入目標單位 ID: ").strip())
    rows=df_t[df_t["ID"]==tid]
    if rows.empty: print(f"{R}找不到 ID {tid}{Z}"); return
    r=rows.iloc[0]
    print(f"\n目前狀態：{sfmt(str(r['狀態']))}  輪次：{r.get('輪次','第一輪')}")

    print("\n更新狀態："); si=pick(STATUSES)
    new_s=STATUSES[si]
    note=input("備註 (可空白): ").strip()
    today=datetime.now().strftime("%Y-%m-%d")

    df_t.loc[df_t["ID"]==tid,"狀態"]=new_s
    df_t.loc[df_t["ID"]==tid,"最後跟進日期"]=today

    # 若回覆了，自動升到第二輪
    if new_s=="已回覆":
        df_t.loc[df_t["ID"]==tid,"輪次"]="第二輪"
        print(f"{C}→ 輪次自動升級為「第二輪」，下次信件將使用正式提案模板{Z}")
        follow=(datetime.now()+timedelta(days=3)).strftime("%Y-%m-%d")
        df_t.loc[df_t["ID"]==tid,"下次跟進日期"]=follow
    elif new_s=="已發信":
        follow=(datetime.now()+timedelta(days=7)).strftime("%Y-%m-%d")
        df_t.loc[df_t["ID"]==tid,"下次跟進日期"]=follow
    elif new_s=="已安排會議":
        mtg=input("會議日期 (YYYY-MM-DD): ").strip()
        df_t.loc[df_t["ID"]==tid,"下次跟進日期"]=mtg

    if note: df_t.loc[df_t["ID"]==tid,"備註"]=note

    log={"日期":today,"單位ID":tid,"單位名稱":r["單位名稱"],
         "聯絡方式":"其他","內容摘要":note or "狀態更新","結果":new_s}
    df_l=pd.concat([df_l,pd.DataFrame([log])],ignore_index=True)
    _save(df_t,df_l)
    print(f"{G}✓ 已更新為：{new_s}{Z}")

# ── 追蹤提醒 ──────────────────────────────────────────────────────────────────
def show_followups():
    df_t=load_t()
    header("追蹤提醒")
    today=datetime.now().date()
    df_t["下次跟進日期"]=pd.to_datetime(df_t["下次跟進日期"],errors="coerce")
    active=df_t[~df_t["狀態"].isin(["合作中","暫不合適"])]

    overdue =active[active["下次跟進日期"].notna()&(active["下次跟進日期"].dt.date<=today)]
    upcoming=active[active["下次跟進日期"].notna()&
                    (active["下次跟進日期"].dt.date>today)&
                    (active["下次跟進日期"].dt.date<=today+timedelta(days=7))]

    if not overdue.empty:
        print(f"\n{R}【逾期跟進】{Z}")
        for _,r in overdue.iterrows():
            d=(today-r["下次跟進日期"].date()).days
            print(f"  ID {int(r['ID'])}  {r['單位名稱']:<20}  {sfmt(r['狀態'])}  {R}逾期{d}天{Z}")

    if not upcoming.empty:
        print(f"\n{Y}【本週待跟進】{Z}")
        for _,r in upcoming.iterrows():
            d=(r["下次跟進日期"].date()-today).days
            print(f"  ID {int(r['ID'])}  {r['單位名稱']:<20}  {sfmt(r['狀態'])}  {Y}{d}天後{Z}")

    if overdue.empty and upcoming.empty:
        print(f"\n{G}  目前沒有待跟進項目{Z}")

    print(f"\n{BOLD}── 整體進度 ──{Z}")
    total=len(df_t)
    for s in STATUSES:
        n=(df_t["狀態"]==s).sum()
        if n: print(f"  {sfmt(s):<20}  {'█'*n} {n}")
    print(f"\n  共 {total} 筆目標單位")
    print(f"\n{BOLD}── 優先序分佈 ──{Z}")
    for p in PRIORITIES:
        n=(df_t["優先序"]==p).sum()
        if n: print(f"  {pfmt(p):<10}  {n} 筆")

# ── 聯絡記錄 ──────────────────────────────────────────────────────────────────
def view_log():
    df_l=load_l()
    header("聯絡記錄")
    if df_l.empty: print(f"{Y}  尚無記錄{Z}"); return
    print(f"\n{'日期':<12}  {'單位名稱':<20}  {'結果':<12}  內容摘要")
    hr()
    for _,r in df_l.tail(30).iterrows():
        print(f"{str(r['日期'])[:10]:<12}  {str(r['單位名稱']):<20}  "
              f"{str(r['結果']):<12}  {str(r['內容摘要'])[:35]}")
    print(f"\n  最近 {min(30,len(df_l))} 筆，共 {len(df_l)} 筆")

# ── PDF 提案書 ────────────────────────────────────────────────────────────────
def gen_pdf_interactive():
    header("生成合作提案書 PDF")
    df_t=load_t()
    view_targets(df_t)
    choice=input("\n輸入目標單位 ID（或直接 Enter 生成通用版）: ").strip()
    if choice:
        rows=df_t[df_t["ID"]==int(choice)]
        if rows.empty: print(f"{R}找不到 ID {choice}{Z}"); return
        r=rows.iloc[0]
    else:
        r={"單位名稱":"","單位類型":""}
    pdf_path=_gen_pdf(r)
    if pdf_path:
        open_yn=input("是否立即開啟 PDF？(y/n): ").strip().lower()
        if open_yn=="y": os.startfile(str(pdf_path))

# ── Email 發送 ────────────────────────────────────────────────────────────────
def load_cfg():
    cfg={}
    # 優先讀取環境變數（Render 雲端部署支援）
    import os as _os
    for key in ["SENDER_EMAIL","SENDER_PASSWORD","SENDER_NAME"]:
        val=_os.environ.get(key)
        if val: cfg[key]=val
    # 本機 config.env 補充（不覆蓋環境變數）
    if CONFIG_FILE.exists():
        for line in CONFIG_FILE.read_text(encoding="utf-8").splitlines():
            line=line.strip()
            if line and not line.startswith("#") and "=" in line:
                k,v=line.split("=",1)
                if k.strip() not in cfg: cfg[k.strip()]=v.strip()
    return cfg

def send_email(to, subj, body, pdf_path=None, quiet=False):
    cfg=load_cfg()
    se=cfg.get("SENDER_EMAIL",""); sp=cfg.get("SENDER_PASSWORD","")
    sn=cfg.get("SENDER_NAME","蒲朝棟")
    if not se or not sp:
        if not quiet:
            print(f"{Y}⚠ 請先設定 config.env（選單 9）{Z}")
        return False
    try:
        msg=MIMEMultipart()
        msg["From"]=f"{sn} <{se}>"
        msg["To"]=to; msg["Subject"]=subj
        msg.attach(MIMEText(body,"plain","utf-8"))
        if pdf_path and Path(pdf_path).exists():
            with open(pdf_path,"rb") as f:
                part=MIMEBase("application","octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f'attachment; filename="{Path(pdf_path).name}"')
            msg.attach(part)
        with smtplib.SMTP_SSL("smtp.gmail.com",465) as s:
            s.login(se,sp); s.sendmail(se,to,msg.as_string())
        if not quiet: print(f"{G}✓ Email 已發送至：{to}{Z}")
        return True
    except Exception as e:
        if not quiet: print(f"{R}✗ 發送失敗：{e}{Z}")
        return False

def email_guide():
    header("Email 發送設定說明")
    print(f"""
  {BOLD}步驟 1｜開啟 Gmail 兩步驟驗證{Z}
    Google 帳戶 → 安全性 → 兩步驟驗證 → 開啟

  {BOLD}步驟 2｜建立應用程式密碼{Z}
    Google 帳戶 → 安全性 → 應用程式密碼
    選擇「郵件」→ 產生 16 碼密碼

  {BOLD}步驟 3｜編輯 config.env{Z}
    {C}SENDER_EMAIL=shoppy09@gmail.com
    SENDER_PASSWORD=（16 碼，無空格）
    SENDER_NAME=蒲朝棟{Z}

  設定完成後，生成信件時可直接發送，批次發信也能自動寄出。
""")

# ── LINE / 電話話術 ───────────────────────────────────────────────────────────
LINE_TMPL = {
    "退輔會/榮服處": "LINE_退輔會.txt",
    "政府勞動局處":  "LINE_政府勞動局.txt",
    "大專院校":     "LINE_大專院校.txt",
    "社區大學":     "LINE_社區大學.txt",
    "企業HR":       "LINE_企業HR.txt",
    "人資協會":     "LINE_人資協會.txt",
}
PHONE_TMPL = {
    "退輔會/榮服處": "電話話術_退輔會.txt",
    "政府勞動局處":  "電話話術_政府勞動局.txt",
    "大專院校":     "電話話術_大專院校.txt",
    "社區大學":     "電話話術_社區大學.txt",
    "企業HR":       "電話話術_企業HR.txt",
    "人資協會":     "電話話術_人資協會.txt",
}

def show_contact_script(script_map, title):
    header(title)
    df_t = load_t()
    # 優先顯示電話先行 or 未聯絡
    subset = df_t[df_t["狀態"]=="未聯絡"]
    view_targets(subset)
    choice = input("\n輸入目標單位 ID: ").strip()
    rows = df_t[df_t["ID"]==int(choice)]
    if rows.empty: print(f"{R}找不到 ID {choice}{Z}"); return
    r = rows.iloc[0]
    fname = script_map.get(str(r["單位類型"]))
    if not fname:
        print(f"{R}找不到對應腳本：{r['單位類型']}{Z}"); return
    path = TEMPLATES_DIR / fname
    if not path.exists():
        print(f"{R}檔案不存在：{fname}{Z}"); return
    import pandas as pd2
    contact = str(r["聯絡人"]) if pd.notna(r["聯絡人"]) and str(r["聯絡人"]).strip() else ""
    text = (path.read_text(encoding="utf-8")
            .replace("{{單位名稱}}", str(r["單位名稱"]))
            .replace("{{聯絡人}}",   contact or "您")
            .replace("{{地區}}",     str(r["地區"]) if pd.notna(r["地區"]) else ""))
    print(f"\n{C}{'─'*60}{Z}")
    print(text)
    print(f"{C}{'─'*60}{Z}")
    # 儲存到 outputs
    out = OUTPUTS_DIR / f"{datetime.now().strftime('%Y%m%d')}_{r['單位名稱']}_{fname}"
    out.write_text(text, encoding="utf-8")
    print(f"{G}✓ 已儲存：outputs/{out.name}{Z}")

# ── A/B 主旨測試 ──────────────────────────────────────────────────────────────
AB_SUBJECTS = {
    "退輔會/榮服處": {
        "A": "職涯轉型課程合作提案｜CDA 國際職涯諮詢師 蒲朝棟",
        "B": "退役軍人的職涯諮詢師，想與貴單位合作｜蒲朝棟",
    },
    "政府勞動局處": {
        "A": "職涯諮詢課程合作提案｜CDA 國際職涯諮詢師 蒲朝棟",
        "B": "想為貴單位求職者提供職涯諮詢服務｜蒲朝棟",
    },
    "大專院校": {
        "A": "職涯講師合作邀請｜CDA 國際職涯諮詢師 蒲朝棟",
        "B": "20年管理實務 + 職涯認證，想為貴校學生提供引導｜蒲朝棟",
    },
    "社區大學": {
        "A": "職涯課程合作提案｜CDA 國際職涯諮詢師 蒲朝棟",
        "B": "「35歲後的轉職策略」課程想與貴校合作｜蒲朝棟",
    },
    "企業HR": {
        "A": "員工職涯發展課程提案｜CDA 國際職涯諮詢師 蒲朝棟",
        "B": "協助員工找到方向，留才比加薪更有效｜蒲朝棟",
    },
    "人資協會": {
        "A": "講師投稿合作邀請｜CDA 國際職涯諮詢師 蒲朝棟",
        "B": "在職諮詢師分享：求職者最常說的3件事｜蒲朝棟",
    },
}

def show_ab_stats():
    header("A/B 主旨測試分析")
    df_t = load_t()
    if "主旨版本" not in df_t.columns:
        print(f"{Y}尚無 A/B 測試資料（發信時會自動記錄版本）{Z}"); return
    sent = df_t[df_t["狀態"].isin(["已發信","已回覆","已安排會議","合作中"])]
    if sent.empty: print(f"{Y}尚無發送記錄{Z}"); return
    print(f"\n{'類型':<14}  {'版本':4}  {'發出':>5}  {'回覆':>5}  {'回覆率':>7}")
    hr()
    for t in TYPES:
        for ver in ["A","B"]:
            sub = sent[(sent["單位類型"]==t) & (sent.get("主旨版本","")==ver)]
            if sub.empty: continue
            total_s = len(sub)
            replied = sub["狀態"].isin(["已回覆","已安排會議","合作中"]).sum()
            rate    = f"{replied/total_s*100:.0f}%" if total_s>0 else "-"
            print(f"  {t:<14}  {ver:4}  {total_s:>5}  {replied:>5}  {rate:>7}")

def _subject_ab(r):
    """A/B 輪流分配主旨"""
    typ    = str(r["單位類型"])
    tid    = int(r["ID"])
    ab     = AB_SUBJECTS.get(typ, {})
    if not ab: return _subject(r)
    ver    = "A" if tid % 2 == 0 else "B"   # 偶數ID用A，奇數用B
    return ab.get(ver, _subject(r)), ver

# ── 週報 ──────────────────────────────────────────────────────────────────────
def weekly_report_interactive():
    header("生成週報")
    from reporter import generate_report, print_weekly_summary
    print_weekly_summary()
    confirm = input("\n是否生成完整 PDF 週報？(y/n): ").strip().lower()
    if confirm=="y":
        generate_report(open_after=True)

# ── 回信偵測 ──────────────────────────────────────────────────────────────────
def check_replies_interactive():
    header("檢查回信")
    from scheduler import check_replies
    check_replies(quiet=False)

# ── 排程發信 ──────────────────────────────────────────────────────────────────
def schedule_interactive_menu():
    header("排程發信管理")
    from scheduler import schedule_interactive, view_queue, process_queue
    print("  1. 新增排程")
    print("  2. 查看現有排程")
    print("  3. 立即執行到期排程")
    c = input("選擇: ").strip()
    if c=="1": schedule_interactive()
    elif c=="2": view_queue()
    elif c=="3": process_queue()

# ── 主選單 ────────────────────────────────────────────────────────────────────
def main():
    ensure_dirs(); init_data()
    while True:
        clr(); header("職涯顧問合作拓展系統  v3.0")
        print(f"\n  {BOLD}蒲朝棟 Tim Pu｜CDA 國際職涯諮詢師｜104 職涯引導師｜桃園{Z}\n")
        print(f"  {BOLD}【基本功能】{Z}")
        print("  1. 查看目標單位")
        print("  2. 新增目標單位")
        print("  3. 生成個人化信件（單筆）")
        print("  4. 批次發信（多筆）")
        print("  5. 更新聯絡狀態")
        print("  6. 追蹤提醒 / 進度統計")
        print()
        print(f"  {BOLD}【聯絡工具】{Z}")
        print("  7. LINE 訊息腳本")
        print("  8. 電話話術腳本")
        print("  9. 排程發信管理")
        print("  R. 檢查回信（自動更新狀態）")
        print()
        print(f"  {BOLD}【報告 & 分析】{Z}")
        print("  W. 生成週報 PDF")
        print("  B. A/B 主旨測試分析")
        print("  L. 聯絡記錄")
        print()
        print(f"  {BOLD}【文件】{Z}")
        print("  P. 生成合作提案書 PDF")
        print("  E. 開啟 Excel 資料庫")
        print("  S. Email 設定說明")
        print("  0. 離開\n")

        c=input("請選擇功能: ").strip().upper()

        if c=="1":
            header("查看目標單位")
            fs=input(f"狀態篩選（Enter跳過）: ").strip() or None
            ft=input(f"類型篩選（Enter跳過）: ").strip() or None
            fr=input("地區篩選（Enter跳過）: ").strip() or None
            fp=input("優先序篩選 [高/中/低]（Enter跳過）: ").strip() or None
            view_targets(fs=fs,ft=ft,fr=fr,fp=fp)
            input("\n按 Enter 繼續...")
        elif c=="2":
            add_target(); input("\n按 Enter 繼續...")
        elif c=="3":
            generate_email_interactive(); input("\n按 Enter 繼續...")
        elif c=="4":
            batch_send(); input("\n按 Enter 繼續...")
        elif c=="5":
            update_status(); input("\n按 Enter 繼續...")
        elif c=="6":
            show_followups(); input("\n按 Enter 繼續...")
        elif c=="7":
            show_contact_script(LINE_TMPL, "LINE 訊息腳本"); input("\n按 Enter 繼續...")
        elif c=="8":
            show_contact_script(PHONE_TMPL, "電話話術腳本"); input("\n按 Enter 繼續...")
        elif c=="9":
            schedule_interactive_menu(); input("\n按 Enter 繼續...")
        elif c=="R":
            check_replies_interactive(); input("\n按 Enter 繼續...")
        elif c=="W":
            weekly_report_interactive(); input("\n按 Enter 繼續...")
        elif c=="B":
            show_ab_stats(); input("\n按 Enter 繼續...")
        elif c=="L":
            view_log(); input("\n按 Enter 繼續...")
        elif c=="P":
            gen_pdf_interactive(); input("\n按 Enter 繼續...")
        elif c=="E":
            if DATA_FILE.exists(): os.startfile(str(DATA_FILE))
        elif c=="S":
            email_guide(); input("\n按 Enter 繼續...")
        elif c=="0":
            print(f"\n{G}再見！祝合作順利！{Z}\n"); break

if __name__=="__main__":
    main()
