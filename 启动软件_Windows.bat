@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo MIMO 工资评分生成器 Final V1.1
echo ============================================
echo.
where python >nul 2>nul
if %errorlevel% neq 0 (
  echo 未检测到 Python。
  echo 新电脑请先安装 Python 3.9 或更高版本，并勾选 Add Python to PATH。
  echo 安装后重新双击本文件。
  echo.
  pause
  exit /b 1
)
python --version
python server.py
pause
