@echo off
chcp 65001 >nul
title 图片文字提取工具 - 一键打包
cd /d "%~dp0"

echo ============================================
echo  图片文字提取工具 打包脚本
echo ============================================
echo.

echo [1/2] 安装依赖...
python -m pip install -r requirements.txt --no-warn-script-location
if errorlevel 1 (
    echo 依赖安装失败，请确认已安装 Python 3.9+ 并勾选 "Add Python to PATH"
    pause
    exit /b 1
)

echo.
echo [2/2] 开始打包（约需 1-3 分钟）...
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "图片文字提取工具" ^
  --icon app_icon.ico ^
  --add-data "app_icon.ico;." ^
  --collect-all rapidocr_onnxruntime ^
  main.py

if errorlevel 1 (
    echo 打包失败，请查看上方错误信息
    pause
    exit /b 1
)

echo.
echo 打包完成！exe 位于: dist\图片文字提取工具.exe
echo 可直接双击运行，也可以复制到本文件夹根目录。
pause
