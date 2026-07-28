# fr-review — 台灣公開發行公司財報實審雙文件產線

給 AI 調度的財報實審工具鏈：輸入**股號＋年度**，產出三份文件——

1. **Excel 查詢函**（「O年度財務比率差異分析說明」）：版面依實審慣例——差異說明
   （損益與成長率、同業比較、資產負債變動、產業發展、風險事項、制式十六問、
   資金貸與背書保證、會計主管）＋基本資料／財報資料／財務比率／成長率各表；
   預填兩期比較數字與變動%，達門檻（30%／10%）自動標註請公司說明，
   「說明／ANS」欄留白給公司填答。
2. **Word 管區意見初稿**（「公開發行○○公司O年度財務報告公告檢查表—管區意見」）：
   18 項檢查表勾選＋五段說明，數字面由程式寫死、質性段落由 AI 依規則填寫，
   公司未回覆前的事項一律標「擬行前查證」。
3. **Word 複核表初稿**（「財務報告實質審閱案件複核表」）：簽核用封面，
   與管區意見同一份內容 JSON 產出。

## 給同事用：一個網址，什麼都不用裝

**<https://shelly-awkward.github.io/fr-review/>**

輸入股票代號 → 下載 Excel 查詢函 → 複製一段文字貼給任何聊天型 AI（ChatGPT／Gemini／
Copilot／Claude 皆可）→ 把 AI 回的內容貼回網頁 → 下載 Word 管區意見與複核表初稿。

網頁在瀏覽器裡直接執行本 repo 的 Python（Pyodide），財報資料與產出**都不會離開使用者電腦**。
內容送出前會自動驗收：數字必須有出處、不得殘留未填欄位、未提供同業平均時不得出現同業平均數值。

## 年度維護（維護者一年做一次，年報 3/31 截止後）

```bash
python scripts/fetch_company_list.py          # 更新公開發行公司名單
python scripts/fetch_archive.py --years 111,112,113,114,115   # 補抓新年度（已有的自動跳過）
python scripts/fetch_archive.py --quarter 2 --years 112,113,114   # 半年報（9 月起抓當年 Q2）
git add data && git commit -m "歸檔 <年度> 年報" && git push
```

`fetch_archive.py` 可隨時中斷再跑，已抓到的會跳過、確定沒申報的會記在
`data/archive_status.json` 不重試。請在台灣 IP 的機器上跑（MOPS 對 GitHub runner IP
段有間歇性封鎖），並保留預設 3 秒間隔，勿高頻打站。

## 快速開始（對 AI 說這句就好）

> 參考 Shelly-awkward/fr-review，幫我寫 8304 佳聯 114年度的實審報告跟請公司說明的文件。

AI 請讀 **[REVIEW_PROMPT.md](REVIEW_PROMPT.md)**（完整四步流程、判斷規則、禁止事項、驗收清單）。

### 對方 AI 沒有終端機怎麼辦

| AI 的能力 | 走哪條路 |
|-----------|----------|
| 有終端機（Gemini CLI、Copilot CLI、Codex CLI、Claude Code、VS Code agent mode） | 直接照 [REVIEW_PROMPT.md](REVIEW_PROMPT.md)，它自己會 clone、裝套件、跑腳本 |
| 有 Python 沙箱但不能連網（ChatGPT 網頁版等） | 你把 repo 下載成 zip 上傳給它，其餘照 REVIEW_PROMPT.md；Word 用 `gen_checklist_docx.py`（沙箱通常沒有 Node） |
| 純聊天、什麼都不能執行 | 走 **[PASTE_PROMPT.md](PASTE_PROMPT.md)**：你在本機跑腳本算數字，只把質性層那一段貼給 AI |

## 架構

```
資料層  data_requests.json ──push──▶ GitHub Actions 抓 MOPS inline-XBRL
        └─ 本機亦可：python scripts/fetch_requests.py（台灣 IP 更穩）
        ▼
        data/<股號>_<西元年>Q<季>_pretrip.json   ← 勤前包（meta/audit/tuples/notes_text/statements/red_flags）

數字層  python scripts/build_review_content.py --co 8304 --year 114 [--quarter 2]
        ▼
        out/<股號>_<年>_inquiry.json          （查詢函題目＋近6年表＋比率表）
        out/<股號>_<年>_review_content.json   （facts 數字＋draft 骨架含【AI待填】佔位）
        （--quarter 2＝半年報：讀 Q2 pretrip、取 1/1–6/30 累計與 6/30 餘額，
        　產出檔名帶 Q2 後綴、文字標籤為「<年>年半年度」；網頁版同步有期別選單）

質性層  AI 依 REVIEW_PROMPT.md 填 draft → out/<股號>_<年>_checklist_content.json
        python scripts/check_content.py …     ← 驗收閘門（佔位字／結構／數字抽核）

產出層  python scripts/gen_inquiry_xlsx.py …  → Excel 查詢函
        python scripts/gen_checklist_docx.py … → Word 管區意見（python-docx 版）
        node   scripts/gen_checklist_docx.js … → Word 管區意見（docx-js 版，輸出相同）
```

**分工鐵律**：Python 只算數字、AI 只填文字，互不越界。同業平均等拿不到的資料
留白標註「請自公開資訊觀測站貼入」——寧缺勿假。

## 為什麼抓取走「改 JSON＋push」

- 雲端 AI session 的網路政策封鎖 MOPS（實測），GitHub App token 通常也無
  workflow_dispatch 權限（403 實測）。
- 但 push 一個 JSON 一定可以：`data_requests.json` 就是抓取佇列，workflow
  `on: push: paths: [data_requests.json]` 讀檔抓取並 commit 回 `data/`。
- 內建「先合併（C）、檔案不存在退個別（A）」fallback——無子公司的公司只申報個別報表。
- 查母公司持股（集團架構）＝把母公司股號的合併報告也加進佇列，同一條路。

## 環境需求

- Python 3.10+，`pip install requests openpyxl python-docx`
- Node 18+（選用）：`cd scripts && npm install`——只有想用 docx-js 版產生器時才需要，
  Word 產出走 Python 版即可，純 Python 環境（含多數 AI 沙箱）也能跑完全程
- 抓取需能連 mopsov.twse.com.tw（台灣 IP 最穩；GitHub Actions 間歇被 WAF 擋，重推即可）

## 範例

`templates/example/example_checklist_content.json`：虛構「典範工業(0000)」的完整成品
內容 JSON，展示金管會體例與兩段式寫法（所有數字虛構，僅供學格式）。
`data/` 內附佳聯(8304)、台數科(6464) 的真實 pretrip（皆屬公開財報資料），可直接當
golden case 跑通全流程。

## 授權與界線

程式碼 MIT。`data/` 內容取自公開資訊觀測站公開申報資料。本工具產出為**初稿**，
請務必自行審核；產出檔（out/）與客戶清單不進版控。
