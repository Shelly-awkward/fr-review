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

### 1-B. 抓完後務必重解析一次　★別漏

抓取程式是 2026-07-27 啟動的，那時 `xbrl_pretrip.py` 還沒擴充附註抽取；Python 在
程式啟動時就把模組載入記憶體，所以**這一輪抓到的檔案都只有舊的 7 個附註概念**
（實測新檔 notes 只有 0–4 節，應為 9–12 節）。原始 HTML 已存在本機，離線重建即可，
不必重抓：

```bash
python scripts/reparse.py
```

跑完 `data/*_pretrip.json` 會全部帶上租賃／金融工具／收入／關聯企業等會計政策原文，
Word 的 IFRS 四段才寫得出「認列及衡量」。**重解析後要再 commit push 一次。**

驗證方式：任取一份新檔，`notes_text` 應有 9 節以上（視公司揭露而定）。

### 2. 依實審表格樣本重排 Excel 版面　✅ 已完成（2026-07-28）

已依 Drive「fr-review交換區／過去範例」三份實際樣本（新永安110、明緯112、南都114Q2）
完成對齊：

- **Excel**（`gen_inquiry_xlsx.py` 重寫）：差異說明（一 損益與成長率、二 同業比較、
  三 資產負債變動、四 產業發展、五 風險事項、六 制式十六問、七 資金貸與及背書保證、
  八 會計主管、連絡人）＋基本資料＋財報資料（7科目×6年）＋財務比率（11比率×6年，
  個別／上市櫃同業平均／所有同業平均）＋成長率（4項×6年）。原「科目分析」表保留為
  最後一張「科目分析(內部)」（內部覆核底稿，寄公司前刪除）。
- **Word 新增複核表**：「財務報告實質審閱案件複核表」由同一份內容 JSON 的 `cover`
  欄位產出（`gen_checklist_docx.py --cover`，網頁第 3 步同時給兩個下載）。
  注意 `gen_checklist_docx.js`（Node 版）**尚未支援 cover**。
- 數字層新增：營業活動現金流量、流動資產、流動負債、普通股股本；
  比率擴為 11 項（流動比率、負債比率、現金流量比率、股東權益報酬率、速動比率
  依 XBRL 計算，口徑可能與個別資料庫查詢系統略異，Excel 已註明）。

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

## 下一階段構想（尚未動工）：公司回覆後的複核

目前產線止於「公司回覆前的初稿」。公司把 Excel 填回來之後，還有一段人工工作：
讀回覆、判斷說明是否合理充分、決定哪些寫進 Word 定稿。這段 AI 可以先行給建議——

- **對答比對**：回覆有沒有正面回答問題（常見的閃避是答非所問、只給結論不給原因）。
- **與數字勾稽**：公司說的原因和財報數字是否一致（例如說「因客戶流失」但應收帳款
  週轉率反而變好）。
- **附件檢核**：承諾要附的評價報告、備查簿、議事錄有沒有真的附上。
- **產定稿**：把採認的回覆內容填進 Word，把「擬行前查證」改寫為查證結果。

一樣走「程式算數字、AI 寫判斷、閘門把關」的分工，且**採認與否的決定權在檢查員**。

## AI 由誰付費

- 預設：同事用自己慣用的 AI（ChatGPT／Gemini／Copilot／Claude 皆可），
  網頁第 2 步按鈕複製 prompt 貼過去即可，維護者不必提供 API key。
- 若日後單位願意編列預算，可在 Apps Script 端加設 API key 集中呼叫，
  同事按一鍵直接拿到兩份完成檔；代價是費用與產出責任歸屬要先講清楚。
- **維護者個人的 API key 不對同事開放。**

## 容量評估（2026-07-27 實測）

單份 pretrip JSON 平均約 60 KB（補附註後估 70 KB），最大 320 KB。

| 保留年度 | 份數 | 容量 |
|----------|------|------|
| 3 年度 | 891 | 約 61 MB |
| 5 年度 | 1,485 | 約 102 MB |
| 10 年度 | 2,970 | 約 203 MB |

GitHub 限制：單檔 100 MB（目前最大 0.32 MB）、repo 建議 1 GB 以內、
Pages 站台 1 GB 以內。**即使保留十年也只用掉五分之一，容量無虞。**
每年新增一個年度約 21 MB。同事查一家公司只下載該公司 3 份 JSON（約 200 KB），
Pages 流量壓力可忽略。

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
