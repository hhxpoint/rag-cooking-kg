#!/bin/bash
# 基于RAG的烹饪知识图谱构建与推理系统 —— 启动脚本（Linux/Mac）
cd "$(dirname "$0")"
echo "正在启动系统..."
python3 src/api/server.py &
sleep 2
xdg-open http://localhost:8000 2>/dev/null || open http://localhost:8000 2>/dev/null || echo "请手动打开 http://localhost:8000"
wait
