@echo off
chcp 65001 >nul
echo ============================================
echo   产品周分析报告 - HTTP 服务器
echo ============================================
echo.
echo 正在启动服务器，请勿关闭此窗口...
echo.
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4"') do set IP=%%a
set IP=%IP:~1%
echo 局域网访问地址: http://%IP%:8080/product-weekly-report.html
echo 本机访问地址:   http://localhost:8080/product-weekly-report.html
echo.
echo 按 Ctrl+C 停止服务器
echo ============================================
echo.
python -m http.server 8080
pause
