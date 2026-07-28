# fr-review：本機 C 槽 ←→ Google Drive 交換區 同步
#
#   sync_to_drive.bat          家裡做完 → 推到 Drive（帶去公司）
#   sync_to_drive.bat back     公司做完 → 收回本機（回家接手）
#   sync_to_drive.bat -WhatIf  只列出會做什麼，不真的動檔案
#
# 接力棒規則：同一時間只有一邊在動工，做完就同步，避免兩邊各改一半。
#
# 不同步（雙向皆排除）：
#   node_modules   重裝即可
#   __pycache__    編譯快取
#   _salvage       本機工作底稿
#   *.html.gz      原始存證檔，太大且可重抓
#   過去範例\      Drive 專屬：實審樣本（版面依據，見 FOLLOW_UP.md）
#   _看我先.md     Drive 專屬：接力說明
#
# ★ /MIR 會刪掉目的地多出來的東西。Drive 上有、本機沒有的資料一定要列進
#   $ExcludeDirs / $ExcludeFiles，否則一次同步就會被清掉（2026-07-28 有前例）。
#
# 主體寫在 .ps1 而非 .bat：cmd 讀批次檔是按位元組定位，檔內有中文時會錯位，
# 導致註解被當指令執行、變數設不進去。sync_to_drive.bat 只是純 ASCII 啟動器。
param(
    [switch]$Back,
    [switch]$WhatIf
)

$ErrorActionPreference = "Stop"
$Local = "C:\Users\dinef\AI\projects\fr-review"
$Drive = "G:\我的雲端硬碟\fr-review交換區"

$ExcludeDirs  = @("node_modules", "__pycache__", "_salvage", "過去範例")
$ExcludeFiles = @("*.html.gz", "_看我先.md")

if ($Back) {
    $src, $dst = $Drive, $Local
    Write-Host "[收回] Google Drive  ->  本機"
} else {
    $src, $dst = $Local, $Drive
    Write-Host "[送出] 本機  ->  Google Drive"
}

if (-not (Test-Path $src)) { Write-Host "** 找不到來源：$src"; exit 1 }
if (-not (Test-Path (Split-Path $dst -Parent))) {
    Write-Host "** 找不到目的地上層目錄，請確認 Google Drive 已掛載。"; exit 1
}

$rcArgs = @($src, $dst, "/MIR", "/XD") + $ExcludeDirs + @("/XF") + $ExcludeFiles +
          @("/R:2", "/W:2", "/NFL", "/NDL", "/NJH")
if ($WhatIf) { $rcArgs += "/L"; Write-Host "（預覽模式：不會真的搬動檔案）" }

& robocopy @rcArgs
$rc = $LASTEXITCODE

# robocopy: 0-7 正常（0=無變更、1=有複製、2=有多餘項、8+=錯誤）
if ($rc -ge 8) {
    Write-Host ""
    Write-Host "** 同步發生錯誤（robocopy 代碼 $rc）**，請確認 Google Drive 已掛載且檔案未被開啟。"
    exit 1
}

Write-Host ""
Write-Host "同步完成（robocopy 代碼 $rc）。提醒："
Write-Host "  - 換一台機器第一次要用時，先在 scripts 資料夾執行 npm install（若要用 .js 版產生器）"
Write-Host "  - 要重建 pretrip（reparse.py）需要 data\*.html.gz，那些沒同步，請重抓"
exit 0
