# fr-review — 台灣公開發行公司財報實審雙文件產線

給 AI 調度的財報實審工具鏈：輸入**股號＋年度**，產出兩份文件——

1. **Excel 查詢函**（「O年度財務報告說明」）：預填兩期比較數字與變動%，依門檻規則
   自動生成給受查公司的問題（30%／10%門檻、減損、OCI、IFRS16/9/15/10、
   資金貸與背書五連問、會計主管四連問），「說明」欄留白給公司填答。
2. **Word 管區意見初稿**（「公開發行○○公司O年度財務報告公告檢查表—管區意見」）：
   18 項檢查表勾選＋五段說明，數字面由程式寫死、質性段落由 AI 依規則填寫，
   公司未回覆前的事項一律標「擬行前查證」。

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

數字層  python scripts/build_review_content.py --co 8304 --year 114
        ▼
        out/<股號>_<年>_inquiry.json          （查詢函題目＋近6年表＋比率表）
        out/<股號>_<年>_review_content.json   （facts 數字＋draft 骨架含【AI待填】佔位）

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
判斷責任在檢查員；產出檔（out/）與客戶清單不進版控。
