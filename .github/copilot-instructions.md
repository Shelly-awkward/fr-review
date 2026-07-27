# AGENTS.md

本 repo 是「台灣公開發行公司財報實審雙文件產線」。使用者要你產出實審文件時：

**先完整讀 [REVIEW_PROMPT.md](REVIEW_PROMPT.md)，照它的四步流程執行。** 那份文件是唯一權威操作說明，
本檔只是指路。

三十秒摘要：

1. 確認 `data/<股號>_<西元年>Q4_pretrip.json` 存在；沒有就把請求加進 `data_requests.json`
   （本機能連 MOPS 就直接 `python scripts/fetch_requests.py`；不能連就 commit＋push 讓
   GitHub Actions 抓，別用 workflow_dispatch）。
2. `python scripts/build_review_content.py --co <股號> --year <民國年>` → `out/` 兩份中介 JSON。
   Excel 查詢函此時即可產出：`python scripts/gen_inquiry_xlsx.py out/<股號>_<年>_inquiry.json <輸出.xlsx>`
3. **你唯一該動筆的地方**：讀 `out/<股號>_<年>_review_content.json`，其中 `facts` 是唯一數字來源、
   `draft` 是含 `【AI待填：…】` 佔位的骨架。把佔位填成正式文字，存成
   `out/<股號>_<年>_checklist_content.json`（只保留 title/groups/footnotes/sections）。
4. `python scripts/check_content.py <content> --review <review>` 必須通過，再產 Word——
   `python scripts/gen_checklist_docx.py <content> <輸出.docx>`（需 python-docx）
   或 `node scripts/gen_checklist_docx.js <content> <輸出.docx>`（需 npm install），
   兩者輸出相同，環境有哪個就用哪個。

鐵律（違反會被 check_content.py 擋下，或更糟——產出不實文件）：

- **不准自己算數字、不准編數字**。只能引用 `facts` 裡有的數字。
- **同業平均等拿不到的資料一律留白並標「行前請至公開資訊觀測站查填」**，寧缺勿假。
- 這是**公司回覆前的初稿**：凡屬公司才答得出的（變動原因、評價假設、議事錄、會計主管進修時數），
  一律寫「擬行前查證／請公司說明」，禁止推測填入。
- 繁體中文、金管會體例、金額新臺幣千元。
- 產出檔一律放 `out/`（已 gitignore）；未經使用者指示不要 commit 或 push 任何東西。

環境：Python 3.10+（`pip install requests openpyxl`）、Node 18+（`cd scripts && npm install`）。
