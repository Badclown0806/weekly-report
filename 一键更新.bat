@echo off
chcp 65001 >nul
cd /d "D:\周汇报文件"

REM 带 openpyxl 的解释器（裸 python 在部分机器上指向无 openpyxl 的运行时，会构建失败）
set "PYTHON=C:\Users\12139\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"

REM Git（优先使用安装路径，缺失时退回 PATH）
set "GIT=C:\Program Files\Git\bin\git.exe"
if not exist "%GIT%" set "GIT=git"

echo ============================================================
echo 一键更新: 生成 HTML(男装+女装) → 提交 → 推送
echo ============================================================

echo.
echo [1/4] 生成 product-weekly-report.html (男装) ...
"%PYTHON%" build_html.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] build_html.py 失败，终止
    pause
    exit /b 1
)

echo.
echo [2/4] 生成 product-weekly-report-women.html (女装) ...
"%PYTHON%" build_html_women.py
if %ERRORLEVEL% neq 0 (
    echo [ERROR] build_html_women.py 失败，终止
    pause
    exit /b 1
)

echo.
echo [3/4] Git add + commit ...
"%GIT%" add product-weekly-report.html data-detail.js product-weekly-report-women.html data_women.js data-detail-women.js
"%GIT%" commit -m "update: data refresh %date%"
if %ERRORLEVEL% neq 0 (
    echo [WARN] git commit 可能无变更或失败，继续尝试 push
)

echo.
echo [4/4] Git push ...
"%GIT%" push origin main
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
