#!/bin/bash
cd "$(dirname "$0")"
echo "============================================"
echo "MIMO 工资评分生成器 Final V1.1"
echo "============================================"
if ! command -v python3 >/dev/null 2>&1; then
  echo "未检测到 Python3。请先安装 Python 3.9 或更高版本。"
  read -p "按回车退出..."
  exit 1
fi
python3 --version
python3 server.py
