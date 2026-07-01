@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ========= 环境检测 =========
where python >nul 2>nul
if %errorlevel% neq 0 (
  echo [失败] 没有检测到 Python。
  echo 需要安装 Python 3.9 或更高版本，并勾选 Add Python to PATH。
  pause
  exit /b 1
)
echo [通过] Python 已安装：
python --version
echo [通过] 软件自带 Excel 依赖 vendor/openpyxl，不需要 pip install。
echo [通过] 保留多种评分策略：严格不超发、批次贴0、逐人四舍五入、多行均衡、最后一场补差。
echo 双击 启动软件_Windows.bat 即可启动。
pause
