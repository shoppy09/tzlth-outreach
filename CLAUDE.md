# 找合作單位系統 - 操作規則

## 系統定位
主動尋找職涯相關合作機會的外展系統。對象包含大專院校就業輔導組、HR 社群、職涯課程機構、講座主辦單位等。

## 角色說明
你是這個 Python 自動化外展系統的操作者與維護者。負責管理目標清單、發送合作邀請信、追蹤回覆狀況、生成報告。

## 技術架構
- 核心程式：main.py（37KB，主要執行邏輯）
- 報告生成：reporter.py + pdf_gen.py
- 排程管理：scheduler.py + 定期發信.py
- Web 介面：web/ 資料夾
- 目標清單：data/targets.xlsx
- 發信模板：templates/（12 種，分人資社群、大專院校、職涯課程機構、講座主辦等，每類有第一封/第二封/LINE 版）
- 產出存放：outputs/（發信稿）、reports/（PDF 報告）

## 目標合作對象類型
1. 人資社群、HR 協會
2. 大專院校就業輔導組
3. 職涯教育課程機構
4. 講座活動主辦單位
5. 企業 HR 部門
6. 多元社群（自媒體相關）

## 目前狀態
已部署至 Render（Singapore）：https://tzlth-outreach.onrender.com ✅ live
本地備份仍可用（web/app.py，需設定環境變數）。

### 雲端架構
- **平台**：Render Free tier（Singapore）
- **Service ID**：srv-d7dnqinlk1mc73esjag0
- **Python 版本**：3.11.9（PYTHON_VERSION env var）
- **系統字體**：fonts-noto-cjk（render.yaml nativePackages）
- **資料持久化**：每次 _save() 觸發背景執行緒 commit targets.xlsx 至 GitHub
- **環境變數**：SENDER_NAME ✅ / GITHUB_TOKEN ✅ / GITHUB_REPO ✅ / SENDER_EMAIL ❌ / SENDER_PASSWORD ❌（Tim 待設定）

## 收尾四件事（每次對話結束前必做）
1. **更新本文件「最近修改記錄」**（見下表）
2. **更新總部任務清單**：`C:\Users\USER\Desktop\tzlth-hq\dev\tasks.md`
3. **更新每日日誌**：`C:\Users\USER\Desktop\tzlth-hq\reports\daily-log.md`
4. **寫入反思日誌**：`C:\Users\USER\Desktop\tzlth-hq\reports\reflection-log.md`（有實質改善價值才寫）

> 未完成收尾四件事 = 任務未完成。

## 最近修改記錄（跨視窗同步用）

| 日期 | 修改內容 | 在哪個視窗執行 | 狀態 |
|------|---------|--------------|------|
| 2026-04-12 | PDF 字體跨平台支援（_FONT_CANDIDATES fallback list + render.yaml nativePackages: fonts-noto-cjk）| 總部視窗 | ✅ live |
| 2026-04-12 | GitHub 資料持久化（github_sync.py + main.py _save() 背景執行緒）| 總部視窗 | ✅ live |
| 2026-04-12 | 部署至 Render（commit 08803bd，解決 Python 版本 + numpy/pandas 相容性）| 總部視窗 | ✅ live |
| 2026-04-12 | 修復看板視圖 404 問題（雙 Flask 進程衝突，kill 舊進程後重啟） | 總部視窗 | ✅ 已確認 |
| 2026-04-12 | base.html 已有「範本管理」導覽連結（/templates），路由 app.py:425 已存在 | 總部視窗 | ✅ 確認可用 |
| 2026-04-18 | 6份第一輪Email模板加入「預診斷洞察」觀察區塊（7行格式）+ 時間戳記更新 | 總部視窗 | ✅ |
| 2026-04-18 | 6份LINE模板加入1句預診斷觀察佔位符 + 時間戳記更新 2026-04-18 | 總部視窗 | ✅ |

## 待辦事項
- [ ] 確認 targets.xlsx 目標清單內容正確
- [ ] 執行首次實際發信測試
- [ ] 確認 scheduler.py 排程設定正確
- [ ] 驗證 outputs/ 與 reports/ 的存檔機制

## 啟動方式
執行 `網頁版.bat`（會開啟 http://localhost:5000）
如遇 404 或路由不對，先執行：
```
netstat -ano | findstr :5000
taskkill /PID [PID號碼] /F
```
再重新執行 bat 檔。

---
## 總部連結（TZLTH-HQ）
- 系統代號：SYS-06
- 總部路徑：C:\Users\USER\Desktop\tzlth-hq
- HQ 角色：業務部的主要執行工具。負責建立外部合作關係、擴大 Tim 的講座與課程機會。
- 存檔規定：每次發信後，outputs/ 自動存檔；每次生成報告後，reports/ 自動存檔（現有機制）
- 拉取欄位：outputs/ 最新檔案（發信進度）、reports/ 最新 PDF（整體狀況）、data/targets.xlsx（目標清單筆數）
---
