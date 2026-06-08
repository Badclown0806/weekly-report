@echo off
chcp 65001 >nul
cd /d "D:\周汇报文件"

echo ============================================================
echo 一键更新: 生成 HTML → 提交 → 推送
echo ============================================================

echo.
echo [1/3] 生成 product-weekly-report.html ...
python build_html.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] build_html.py 失败，终止
    pause
    exit /b 1
)

echo.
echo [2/3] Git add + commit ...
"C:\Program Files\Git\bin\git.exe" add product-weekly-report.html data-detail.js
"C:\Program Files\Git\bin\git.exe" commit -m "update: data refresh %date%"
if %ERRORLEVEL% neq 0 (
    echo [WARN] git commit 可能无变更或失败，继续尝试 push
)

echo.
echo [3/3] Git push ...
"C:\Program Files\Git\bin\git.exe" push origin main
if %ERRORLEVEL% neq 0 (
    echo [ERROR] git push 失败
    pause
    exit /b 1
)

echo.
echo ============================================================
echo 完成!
echo ============================================================
pause