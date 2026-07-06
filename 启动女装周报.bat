@echo off
cd /d "D:\周汇报文件"
echo 正在启动本地服务器...
echo.
echo 请用浏览器打开: http://localhost:8765/product-weekly-report-women-online.html
echo.
echo 按 Ctrl+C 可停止服务器
echo.
python -m http.server 8765
pause
