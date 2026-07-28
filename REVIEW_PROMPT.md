# REVIEW_PROMPT.md — 財報實審雙文件產線（寫給任何 AI 的操作說明）

你是一個 AI 助手，使用者要你「參考這個 repo，幫 XX 公司（股號）寫 O 年度的實審文件」。
照本文四步流程執行，直出三份文件：

1. **Excel 查詢函**（「O年度財務比率差異分析說明」）——給受查公司填答，
   版面依實審慣例（差異說明一～八節＋基本資料／財報資料／財務比率／成長率），
   數字由數字層自動預填。
2. **Word 管區意見初稿**（「公開發行○○公司O年度財務報告公告檢查表—管區意見」）——
   公司回覆前就先寫好的版本，質性段落由你（AI）依規則填寫。
3. **Word 複核表初稿**（「財務報告實質審閱案件複核表」）——簽核用封面，
   與管區意見同一份內容 JSON 的 `cover` 欄位，產生器自動另存一份 docx。

## 輸入

| 項目 | 必要性 | 說明 |
|------|--------|------|
| 股號、年度 | 必要 | 年度＝民國或西元皆可（114＝2025） |
| 前次實審文件 | 選填 | 有則沿用其風險事項脈絡與寫法 |
| 公司「資金貸與及背書保證作業程序」PDF | 選填 | 有則核對內規限額（訂定面） |
| 同業平均數 JSON | 選填 | 格式見 `scripts/build_review_content.py` 檔頭；沒有就留白 |
| 母公司／同業股號 | 選填 | 需查集團持股時，把母公司合併報告也加進抓取佇列 |

## 四步流程

### ① 備資料（data/ 沒有該公司該年度的 pretrip 時才需要）

檢查 `data/<股號>_<西元年>Q4_pretrip.json` 是否存在（半年報實審為 `Q2`，
season 填 2）。**近6年表最好有 3 份年報**
（每份含兩年數字：目標年、目標-2、目標-4）。缺的話：

- **雲端 AI（連不到 MOPS）**：把請求加進 `data_requests.json` 並 commit＋push——
  例 `{"co": "8304", "year": 2025, "season": 4}`——GitHub Actions 會自動抓取並
  commit 進 `data/`，等（約 3–10 分鐘）後 pull。失敗（MOPS WAF 間歇擋 GitHub IP）
  就再 push 一次重觸發。**不要用 workflow_dispatch**（多數 AI token 無此權限，403）。
- **本機（台灣 IP）**：直接 `python scripts/fetch_requests.py`。

抓取內建「先合併（C）、檔案不存在再退個別（A）」——無子公司的公司只申報個別報表。

### ② 跑數字層（純 Python，你不要自己算）

```bash
python scripts/build_review_content.py --co 8304 --year 114 [--quarter 2] [--peer-avg peers.json]
```

產出 `out/<股號>_<民國年>_inquiry.json` 與 `out/<股號>_<民國年>_review_content.json`。
半年報（`--quarter 2`）讀 Q2 pretrip、取上半年累計數與 6/30 餘額，
產出檔名帶 `Q2` 後綴（如 `8304_114Q2_inquiry.json`），文字標籤自動為「114年半年度」。
所有變動%、週轉率、限額核對都在裡面，**你只准引用，不准重算、不准編**。

Excel 這時就能直接出（查詢函不需要 AI 填寫，「說明」欄本來就留白給公司）：

```bash
python scripts/gen_inquiry_xlsx.py out/8304_114_inquiry.json "out/8304_114_財務比率差異分析說明.xlsx"
```

### ③ 你（AI）填質性段落

讀 `review_content.json`：`facts` 是唯一數字來源，`draft` 是含
`【AI待填：…指引…】` 佔位的骨架。把 draft 填成最終內容，存成
`out/<股號>_<民國年>_checklist_content.json`（只留 title/cover/groups/footnotes/sections，
去掉 meta/facts 外殼）。`cover` 是複核表欄位：公司背景介紹依 facts 既有資訊撰寫
（查不到的寫「請自行查填」）、風險事項與 sections 一、一致、複核意見／結論用標準語
（無異常時「尚無重大異常」「尚無發現重大異常，文擬陳閱後存查，當否？謹請核示。」）。
體例範本：`templates/example/example_checklist_content.json`
（虛構公司，學格式不抄內容）。

**判斷規則：**

- 金管會體例、繁體中文、金額新臺幣千元；口吻「經查…核尚無重大異常」「無此情事」。
- 檢查表備註簡短：「詳個別資料庫說明」「無此情事」「尚無重大異常」；每項 mark 填
  `"yes"`（正常）或 `"no"`（異常，備註必須寫明事由）。
- **初稿定位**：公司尚未回覆。凡屬公司才答得出的（變動原因細節、評價假設、議事錄內容、
  會計主管進修時數），一律寫「請自行確認」「請洽公司說明」——**禁止推測填入**。
  數字面（變動%、勾稽、限額核對）則直接寫死，不加「約」「大概」。
- 三、財務報告審閱說明四項（IFRS16／IFRS9+IFRS7／IFRS15／IFRS10 B38~B50）每項採
  「認列及衡量：…／財報揭露：…」兩段式；會計政策內容依 `facts` 與財報附註
  （`notes_text`、pretrip 有的部分）寫實，pretrip 沒有的政策細節寫「請自行核閱財報附註確認」。
- 風險事項段：draft 給的是「（候選風險）」，你要判斷取捨、改成「（一）（二）…」正式編號；
  不成立的候選整段刪除；每項風險寫明「因應措施」或「請自行確認」。
- 表格段落以 `【表】a｜b｜c` 一列一段，產生器自動轉表格。
- 末段「附註：資料來源與限制」誠實列出實際用到的來源；沒用到的不要列。

**禁止事項：**

- 禁止引用 `facts` 沒有的數字（同業平均拿不到就留白標「請自行至公開資訊觀測站查填」）。
- 禁止留任何 `【AI待填`、`（候選風險）`、`待補`、`TODO` 字樣。
- 禁止把範例（典範工業）的內容抄進正式文件。

### ④ 驗收＋產 Word

```bash
python scripts/check_content.py out/8304_114_checklist_content.json --review out/8304_114_review_content.json
```

閘門過了才產 Word。兩支產生器輸出**完全相同**，看你的環境有什麼就用哪支：

```bash
# 有 Python（pip install python-docx）——沒有 Node 的環境用這支
# 內容 JSON 有 cover 時會自動另存複核表 docx（--cover 可指定路徑）
python scripts/gen_checklist_docx.py out/8304_114_checklist_content.json "out/佳聯有線電視(8304)114年度財務報告實審_公告檢查表-管區意見.docx"
```

```bash
# 有 Node（cd scripts && npm install）——注意 .js 版尚未支援複核表（cover）
node scripts/gen_checklist_docx.js out/8304_114_checklist_content.json "out/佳聯有線電視(8304)114年度財務報告實審_公告檢查表-管區意見.docx"
```

`check_content.py` 未過（退出碼 1）就回步驟③修，**不得跳過閘門**。

## 驗收清單（最終自查）

- [ ] 內容 JSON 所有金額與 `facts` 一致（check_content.py 已抽核，仍請自查衍生計算）
- [ ] 無佔位字、無「待補」
- [ ] 不可得資訊（同業平均、進修時數、議事錄、作業程序未提供時）都標「請自行確認」，沒有瞎編
- [ ] Excel 版面：差異說明一～八節（損益與成長率／同業比較／資產負債變動／產業發展／
      風險事項／制式十六問／資金貸與背書保證／會計主管）＋基本資料／財報資料／財務比率／成長率
- [ ] Word 結構：檢查表（勾選＋備註）＋五段說明＋資料來源與限制；複核表（cover）欄位填妥
- [ ] 渲染目檢：有 LibreOffice 時 `soffice --headless --convert-to pdf <docx>` 轉 PDF 檢視
      （雲端未裝可 `apt-get install -y libreoffice-writer poppler-utils`）；沒有就至少
      重讀一遍 JSON 內容

## 紀律

- 客戶清單、批次選案產出**不進版控**；單一公司實審成品是否 commit 由使用者決定。
- 產出檔一律放 `out/`（已 gitignore）。
- 這是公開 repo：commit 進來的東西人人可見，機密內容（內部意見定稿、客戶名單）不要推上來。
