@echo off
chcp 65001 >nul
REM ─────────────────────────────────────────────────────────────
REM fr-review：本機 C 槽 ←→ Google Drive 交換區 同步
REM
REM   sync_to_drive.bat          家裡做完 → 推到 Drive（帶去公司）
REM   sync_to_drive.bat back     公司做完 → 收回本機（回家接手）
REM
REM 接力棒規則：同一時間只有一邊在動工，做完就同步，避免兩邊各改一半。
REM 不同步的東西：node_modules（重裝即可）、data\*.html.gz（原始存證檔，
REM 太大且可重抓）、_salvage（本機工作底稿）。
REM ─────────────────────────────────────────────────────────────
setlocal
set "LOCAL=C:\Users\dinef\AI\projects\fr-review"
set "DRIVE=G:\我的雲端硬碟\fr-review交換區"

if /I "%~1"=="back" (
  set "SRC=%DRIVE%"
  set "DST=%LOCAL%"
  echo [收回] Google Drive  →  本機
) else (
  set "SRC=%LOCAL%"
  set "DST=%DRIVE%"
  echo [送出] 本機  →  Google Drive
)

robocopy "%SRC%" "%DST%" /MIR /XD node_modules __pycache__ _salvage /XF *.html.gz /R:2 /W:2 /NFL /NDL /NJH
set RC=%ERRORLEVEL%
if %RC% GEQ 8 (
  echo.
  echo ** 同步發生錯誤（robocopy 代碼 %RC%）**，請確認 Google Drive 已掛載且檔案未被開啟。
  exit /b 1
)
echo.
echo 同步完成。提醒：
echo   - 換一台機器第一次要用時，先在 scripts 資料夾執行 npm install（若要用 .js 版產生器）
echo   - 要重建 pretrip（reparse.py）需要 data\*.html.gz，那些沒同步，請重抓
endlocal
