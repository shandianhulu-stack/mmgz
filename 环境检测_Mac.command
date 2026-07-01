#!/bin/bash
cd "$(dirname "$0")"
echo "========= 环境检测 ========="
if ! command -v python3 >/dev/null 2>&1; then
  echo "[失败] 没有检测到 Python3。需要安装 Python 3.9 或更高版本。"
  read -p "按回车退出..."
  exit 1
fi
echo "[通过] Python 已安装："
python3 --version
echo "[通过] 软件自带 Excel 依赖 vendor/openpyxl，不需要 pip install。"
echo "[通过] 保留多种评分策略：严格不超发、批次贴0、逐人四舍五入、多行均衡、最后一场补差。"
echo "双击 启动软件_Mac.command 即可启动。"
read -p "按回车退出..."
