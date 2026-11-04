@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo === PLEX 每日商机：正在抓取并生成仪表盘 ===
python scripts\run_daily.py
if errorlevel 1 (
  echo.
  echo [出错] 看 logs\ 目录下的日志
  pause
  exit /b 1
)
echo.
echo 完成。正在打开仪表盘...
start "" "dashboard.html"
