@echo off
REM 基于RAG的烹饪知识图谱构建与推理系统 —— 启动脚本（Windows）
cd /d %~dp0
echo 正在启动系统，请稍候...
start "" http://localhost:8000
python src\api\server.py
pause
