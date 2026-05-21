@echo off
chcp 950 > nul
title 自動打包 + 同步至 dist-0
setlocal

echo.
echo ============================================================
echo  -- 0.bat: PyInstaller 打包 + 同步到 dist-0
echo ============================================================
echo  功能：執行 pyinstaller 重新打包，並把 dist 同步到 dist-0
echo        （增量同步，不會刪除 dist-0 中的個人檔案）
echo ============================================================

set "PROJ=D:\Dropbox\projects\A"
set "SRC=%PROJ%\dist\地籍資料查詢系統"
set "DST=%PROJ%\dist-0\地籍資料查詢系統"

cd /d "%PROJ%"

echo.
echo ============================================================
echo  [1/2] PyInstaller 打包中...
echo ============================================================
echo.

pyinstaller "地籍資料查詢系統.spec" -y
if errorlevel 1 (
    echo.
    echo ============================================================
    echo  [X] 打包失敗，腳本中止
    echo ============================================================
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  [2/2] 同步到 dist-0（只新增/更新，不刪除目的端的檔案）
echo ============================================================
echo  來源：%SRC%
echo  目的：%DST%
echo.
echo  正在同步檔案，請稍候...

rem 注意：用 /E 不用 /MIR，避免刪掉 dist-0 內的使用者資料
rem （如 data.json、python_embedded、工作資料夾等）
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
echo  [OK] 全部完成！可以開始測試 dist-0 內的 exe
echo ============================================================
echo.
echo 按任意鍵關閉視窗...
pause > nul
