@echo off
chcp 950 > nul
title 同步 dist → dist-0
setlocal

echo.
echo ============================================================
echo  -- 1.bat: 只同步 dist 到 dist-0（不重新打包）
echo ============================================================
echo  功能：當你已手動跑過 pyinstaller，只想把 dist 內容
echo        同步到 dist-0 時用此檔（比 0.bat 快）
echo ============================================================

set "PROJ=D:\Dropbox\projects\A"
set "SRC=%PROJ%\dist\地籍資料查詢系統"
set "DST=%PROJ%\dist-0\地籍資料查詢系統"

echo.
echo ============================================================
echo  同步到 dist-0（只新增/更新，不刪除目的端的檔案）
echo ============================================================
echo  來源：%SRC%
echo  目的：%DST%
echo.
echo  正在同步檔案，請稍候...

rem 注意：用 /E 不用 /MIR，避免刪掉 dist-0 內的使用者資料
robocopy "%SRC%" "%DST%" /E /MT:16 /NFL /NDL /NJH /NJS /R:1 /W:1 > nul
set "RC=%ERRORLEVEL%"

if %RC% GEQ 8 (
    echo.
    echo ============================================================
    echo  [X] 同步失敗 ^(robocopy exit code %RC%^)
    echo ============================================================
    pause
    exit /b 1
)

powershell -Command "[System.Console]::Beep(800, 200); [System.Console]::Beep(1200, 200)" 2>nul

echo  同步完成。
echo.
echo ============================================================
echo  [OK] 完成！可以開始測試 dist-0 內的 exe
echo ============================================================
echo.
echo 按任意鍵關閉視窗...
pause > nul
