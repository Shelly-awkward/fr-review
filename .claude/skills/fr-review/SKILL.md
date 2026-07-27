---
name: fr-review
description: 財報實審雙文件產線。使用者說「幫我寫 XXXX（股號）O年度的實審報告」「產實審文件」「寫管區意見」「出查詢函／財務報告說明」時使用。流程＝抓 MOPS pretrip → Python 數字層出中介 JSON → AI 填質性段落 → 產 Excel 查詢函＋Word 管區意見初稿。
---

# fr-review：財報實審雙文件產線

完整操作說明在 repo 根目錄 **REVIEW_PROMPT.md**——先讀它，照四步流程走。
本檔只補 Claude Code 環境的執行細節。

## 執行摘要

1. **備資料**：查 `data/<股號>_<西元年>Q4_pretrip.json`。缺→本機跑
   `python scripts/fetch_requests.py`（先把請求加進 `data_requests.json`）；
   雲端→改 `data_requests.json` push 等 Actions。近6年表需 3 份年報（隔年抓：Y、Y-2、Y-4）。
2. **數字層**：`python scripts/build_review_content.py --co <股號> --year <民國年>`
   → `out/` 兩份中介 JSON。Excel 可即出：`python scripts/gen_inquiry_xlsx.py …`。
3. **質性層**（本 AI 的工作）：讀 `review_content.json`，`facts` 為唯一數字來源，
   把 `draft` 的 `【AI待填：…】` 填成正式文字，存 `out/<股號>_<年>_checklist_content.json`。
   規則與禁止事項詳 REVIEW_PROMPT.md 步驟③——重點：初稿不推測，公司才答得出的寫
   「擬行前查證」；不引用 facts 沒有的數字。
4. **驗收＋產出**：`python scripts/check_content.py <content> --review <review>` 過閘門
   （未過必修，不得跳過）→ `node scripts/gen_checklist_docx.js <content> <out.docx>`。

## 本機注意

- Windows 中文輸出：腳本已內建 `sys.stdout.reconfigure(encoding="utf-8")`。
- `scripts/node_modules` 不在則先 `cd scripts && npm install`（docx ^9.7.1）。
- 使用者若提供公司作業程序 PDF／前次實審 doc／同業平均數，依 REVIEW_PROMPT.md
  輸入表運用；同業平均做成 JSON 用 `--peer-avg` 餵進數字層。
- 產出檔一律放 `out/`（gitignored）；要不要 commit 成品由使用者決定。
