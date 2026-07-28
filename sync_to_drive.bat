@echo off
REM fr-review <-> Google Drive sync. Logic lives in sync_to_drive.ps1
REM (cmd mis-parses batch files containing non-ASCII text, so keep this file ASCII-only).
REM
REM   sync_to_drive.bat            local  -> Drive
REM   sync_to_drive.bat back       Drive  -> local
REM   sync_to_drive.bat -WhatIf    preview only, no file changes
setlocal
set "PSARGS="
if /I "%~1"=="back" set "PSARGS=-Back"
if /I "%~1"=="-WhatIf" set "PSARGS=-WhatIf"
if /I "%~2"=="-WhatIf" set "PSARGS=%PSARGS% -WhatIf"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync_to_drive.ps1" %PSARGS%
exit /b %ERRORLEVEL%
