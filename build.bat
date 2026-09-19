@echo off
chcp 65001 >nul
title 图片文字提取工具 - 一键打包
cd /d "%~dp0"

echo ============================================
echo  图片文字提取工具 打包脚本（轻量化 onedir 版）
echo ============================================
echo.

echo [1/3] 安装依赖...
python -m pip install -r requirements.txt --no-warn-script-location
if errorlevel 1 (
    echo 依赖安装失败，请确认已安装 Python 3.9+ 并勾选 "Add Python to PATH"
    pause
    exit /b 1
)

echo.
echo [2/3] 开始打包（约需 1-3 分钟）...
python -m PyInstaller --noconfirm --clean --onedir --windowed ^
  --name "图片文字提取工具" ^
  --icon app_icon.ico ^
  --add-data "app_icon.ico;." ^
  --collect-all rapidocr_onnxruntime ^
  --collect-all openpyxl ^
  --exclude-module matplotlib ^
  --exclude-module scipy ^
  --exclude-module pandas ^
  --exclude-module IPython ^
  --exclude-module pytest ^
  --exclude-module tkinter.test ^
  --exclude-module pydoc ^
  --exclude-module distutils ^
  main.py

if errorlevel 1 (
    echo 打包失败，请查看上方错误信息
    pause
    exit /b 1
)

echo.
echo [3/3] 精简体积（移除不需要的 FFmpeg 和元数据）...
del /q "dist\图片文字提取工具\_internal\cv2\opencv_videoio_ffmpeg500_64.dll" 2>nul
for /d %%d in ("dist\图片文字提取工具\_internal\*.dist-info") do rmdir /s /q "%%d" 2>nul

echo.
echo ============================================
echo  打包完成！
echo  程序位于: dist\图片文字提取工具\图片文字提取工具.exe
echo  启动速度快，无需解压，低配电脑也能流畅运行
echo ============================================
pause
