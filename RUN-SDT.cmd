@echo off
setlocal
REM ============================================================================
REM  MAGNA5 SDT - offline / air-gapped launcher
REM
REM  Double-click this file. It installs SDT from the files sitting next to it,
REM  with zero internet access required: portable Python, plink and the vSphere
REM  packages all ship in deps\.
REM
REM  Why a .cmd and not the .ps1 directly: PowerShell's execution policy and
REM  Mark-of-the-Web both block scripts that arrived inside a downloaded zip.
REM  This wrapper clears both, for this one process only.
REM ============================================================================
cd /d "%~dp0"

if not exist "%~dp0install.ps1" (
    echo.
    echo   [X] install.ps1 not found next to this file.
    echo       Extract the whole zip first, then run RUN-SDT.cmd from inside it.
    echo.
    pause
    exit /b 1
)

echo.
echo   Magna5 SDT - offline install
echo   ----------------------------
echo   Source : %~dp0
echo.

REM Strip Mark-of-the-Web from everything in the bundle before anything runs.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%~dp0.' -Recurse -File -EA 0 | Unblock-File -EA 0" >nul 2>&1

set SDT_OFFLINE=1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1" -LocalSource "%~dp0." %*
set RC=%ERRORLEVEL%

echo.
if not "%RC%"=="0" (
    echo   [X] Install exited with code %RC%.
    echo       Read the messages above - the failing step names itself.
) else (
    echo   Done. Open a NEW terminal and type:  sdt
)
echo.
pause
endlocal
exit /b %RC%
