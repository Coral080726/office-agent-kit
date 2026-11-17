@echo off
chcp 65001 >nul
echo 正在启动 OfficeAgent Kit 本地网页版 -> http://127.0.0.1:8766
python office_agent.py webui --port 8766
pause
