# FOLLOW UP — 2026-07-27 收工狀態與待辦

接手者（人或 AI）請從這裡開始。專案全貌看 [README.md](README.md)，
AI 操作流程看 [REVIEW_PROMPT.md](REVIEW_PROMPT.md)。

## 現況

| 項目 | 狀態 |
|------|------|
| 公開網站 | <https://shelly-awkward.github.io/fr-review/> 已上線可用 |
| GitHub | <https://github.com/Shelly-awkward/fr-review>（public） |
| 資料歸檔 | 297 家公開發行公司 × 110–114 年度，**抓取中**（見待辦 1） |
| 產線 | Excel 查詢函＋Word 管區意見初稿，8304 佳聯 golden case 全程驗證通過 |
| 跨 AI | Gemini／Copilot／ChatGPT／Claude 皆可（指令檔 AGENTS.md／GEMINI.md／.github/copilot-instructions.md） |

## 待辦（依重要性）

### 1. 抓取尚未跑完　★最優先

2026-07-27 晚間啟動全量抓取，尚未跑完即收工。**接手第一件事就是續跑**：

```bash
python scripts/fetch_archive.py --years 110,111,112,113,114
```

已抓到的會自動跳過，確定沒申報的記在 `data/archive_status.json` 不重試，
所以隨時中斷、隨時續跑都安全。跑完務必：

```bash
git add data && git commit -m "全量歸檔 297 家 110-114 年度" && git push
```

沒 push 之前，網站上只看得到已 commit 的公司。

### 2. 依實審表格樣本重排 Excel 版面

使用者要提供 3–4 份不同公司的實際實審 Excel。拿到後要對齊的是：
科目清單（是否要列到更細的項目）、比較期間（幾期）、版面與表頭格式。
目前的「科目分析」表是依分析邏輯自行設計的，**不是**依既有作業慣例。

### 3. Gmail 寄送尚未接通

`mailer/Code.gs` 已寫好但**還沒部署**，網頁也還沒接上。剩下的環節：

1. 開 <https://script.google.com> → 新增專案 → 貼上 `mailer/Code.gs` 全文。
2. 專案設定 → 指令碼屬性 → 新增 `PASSWORD`（發給同事的密碼）與
   `RECIPIENT`（收件 Gmail，多個以逗號分隔）。
3. 部署 → 新增部署作業 → 網頁應用程式 → 執行身分「我」、存取權「任何人」。
4. 複製 `https://script.google.com/macros/s/…/exec` 網址。
5. **回到 index.html 加寄送 UI**：第 3 步下方加「密碼」欄與「寄到信箱」按鈕，
   把產出的 xlsx／docx 轉 base64 後 POST 給該網址，body 格式見 Code.gs 的 doPost。
6. 測試：密碼錯要擋、附件要收得到、Gmail 每日寄信配額約 100 封。

密碼只在 Apps Script 端驗證——**不要**把密碼寫進 index.html，那等於公開。

### 4. 頁碼補填

體例用**頁碼**（非附註編號）。XBRL 沒有頁碼，所以 Word 初稿在 IFRS 各段句末留
「（詳財務報告第＿頁）」標記，`check_content.py` 會在通過訊息後回報有幾處待補，
由承辦對照財務報告填入。若之後要自動化，得從財報 PDF 取頁碼，目前沒做。

### 5. 資料保留規則（每年 update 時務必注意）

**成長率的「兩期比較」需要三個年度的數字**（本期成長率 vs 上期成長率），
所以每家公司至少要保留**連續三個年度**的年報。年度歸檔時只新增、不要刪舊檔。
`build_review_content.py` 會自動檢查並在資料不足時印警告、Excel 首列也會標示，
但預防勝於補救。

## 在公司電腦接手（辦公室擋 GitHub）

Google Drive 交換區：`G:／我的雲端硬碟／fr-review交換區`
（規則見該資料夾的 `_看我先.md`）。

- 家裡做完 → 執行 `sync_to_drive.bat`（送出）
- 公司做完 → 回家執行 `sync_to_drive.bat back`（收回）
- **接力棒規則：同一時間只有一邊動工。** 兩邊各自 commit 會讓 git 歷史分岔。
- 未同步的東西：`data` 內的 `.html.gz` 原始存證檔（可重抓）、`node_modules`
  （公司端改用 Python 版 Word 產生器）、`_salvage`。
- ⚠ 2026-07-27 複製當下抓取仍在進行，Drive 上的 `data` 不是最終版；
  抓取跑完後請再執行一次 `sync_to_drive.bat`。

## 明年 4 月的年度作業

1. GitHub → Actions → 「年度歸檔」→ Run workflow（或本機跑 `fetch_archive.py`）。
2. MOPS 會間歇性封鎖 GitHub 的 IP，雲端失敗率高時改在**台灣 IP 的本機**補跑。
3. 抓完 commit push，網站自動更新（`data/index.json` 由腳本自動重建）。

## 已知限制

- 同業平均：MOPS 財務業務資訊未自動抓取，Excel 留白、Word 標「行前查填」。
  程式會**主動擋下**任何憑空填入的同業平均數值。
- 金融保險業（58 家）、證券業（24 家）的 XBRL 科目表與一般產業不同，
  抓不到的科目一律留白不編造（實測亞東證券可正常產出，只是科目較少）。
- 財報附註只抽了 9 個關鍵節（租賃、金融工具、收入、關聯企業、減損、不動產等），
  非全文；其餘政策細節 Word 會寫「擬行前核閱財報附註確認」。
- 原始存證檔不進 git。換一台電腦要用 `reparse.py` 重建 pretrip 時，得先重抓原始檔。

## 檔案地圖

```
index.html            瀏覽器版（Pyodide 跑 scripts 內的 .py）
REVIEW_PROMPT.md      給 AI 的四步流程與判斷規則（權威文件）
PASTE_PROMPT.md       給「不能執行程式」的純聊天 AI 用的可貼上版
AGENTS.md / GEMINI.md / .github/copilot-instructions.md   各家 AI 的自動發現入口
sync_to_drive.bat     與 Google Drive 交換區雙向同步
scripts/
  fetch_company_list.py   MOPS 公發公司名單 → data/companies.json
  fetch_archive.py        全量抓取（可中斷續跑）
  fetch_requests.py       單筆／少量抓取（data_requests.json 佇列）
  reparse.py              用本機原始存證檔離線重建 pretrip（改解析規則後用）
  build_review_content.py 數字層：pretrip → inquiry.json + review_content.json
  gen_inquiry_xlsx.py     Excel 查詢函（含科目分析總表）
  gen_checklist_docx.py   Word 管區意見（python-docx 版）
  gen_checklist_docx.js   Word 管區意見（docx-js 版，輸出相同）
  check_content.py        驗收閘門（佔位字／結構／數字出處／同業平均造假）
mailer/Code.gs        Gmail 寄送端點（待部署）
data/                 pretrip JSON 歸檔＋companies.json＋index.json
```
