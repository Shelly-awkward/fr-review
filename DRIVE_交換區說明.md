# fr-review 交換區（Google Drive）

這是 `C:\Users\dinef\AI\projects\fr-review` 的同步副本，供**辦公室電腦**接手作業用
（辦公室擋 GitHub，故走 Drive，規則同 `listdata交換區` 的接力棒模式）。

## 接力棒規則

**同一時間只有一邊動工，做完就同步。** 兩邊同時改會產生衝突，Drive 不會幫你合併。

| 情境 | 在哪台跑 | 指令 |
|------|----------|------|
| 家裡做完，要帶去公司 | 家裡 | `sync_to_drive.bat` |
| 公司做完，回家接手 | 家裡 | `sync_to_drive.bat back` |

（`sync_to_drive.bat` 在本機專案資料夾內；公司端直接改這裡的檔案即可，
Drive 會自動同步回雲端。）

## 在公司可以做什麼

需要 Python 3.10+ 與 `pip install openpyxl python-docx`（要抓 MOPS 另需 `requests`）。

```
python scripts\build_review_content.py --co 8304 --year 114
python scripts\gen_inquiry_xlsx.py out\8304_114_inquiry.json out\查詢函.xlsx
python scripts\check_content.py out\8304_114_checklist_content.json --review out\8304_114_review_content.json
python scripts\gen_checklist_docx.py out\8304_114_checklist_content.json out\管區意見.docx
```

也可以直接雙擊 `index.html`？**不行**——網頁要用 http 開啟才能載入資料，
請在這個資料夾開命令列執行 `python -m http.server 8765`，再開瀏覽器到
<http://localhost:8765>。或直接用線上版 <https://shelly-awkward.github.io/fr-review/>
（公司若也擋 github.io 就走本機這條）。

## 沒有同步過來的東西

- `data\*.html.gz`：MOPS 原始存證檔，體積大且可重抓。只有要用 `reparse.py`
  重建 pretrip 時才需要，公司端用不到。
- `scripts\node_modules`：Word 產生器的 Node 版依賴。公司端請改用
  `gen_checklist_docx.py`（Python 版，輸出相同），不必裝 Node。
- `_salvage\`：本機工作底稿，不外流。

## 接手看哪份文件

1. **FOLLOW_UP.md** ← 從這裡開始，寫了收工狀態與五項待辦
2. README.md ← 專案全貌
3. REVIEW_PROMPT.md ← 要 AI 幫忙時給它讀這份

## 注意

`.git` 有一併同步，所以這裡是完整的 git repo，可以在公司 commit（純本機，
不需要連 GitHub）。回家同步後再 push 即可。**但請務必遵守接力棒規則**——
兩邊各自 commit 會讓歷史分岔，處理起來很麻煩。
