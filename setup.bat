@echo off
chcp 65001 >nul

:: IMS Print Windows 安裝腳本
echo 🚀 開始安裝 IMS Print 調貨單生成器...

:: 檢查 Python 是否安裝
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 請先安裝 Python 3.7 或更高版本
    echo 💡 下載地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 顯示 Python 版本
echo 📋 檢測 Python 版本:
python --version

:: 檢查 tkinter 是否可用
echo 🔍 檢查 tkinter 模組...
python -c "import tkinter; print('✅ tkinter 模組可用')" 2>nul
if %errorlevel% neq 0 (
    echo ❌ tkinter 模組不可用
    echo 💡 請重新安裝 Python 並確保勾選 "tcl/tk and IDLE" 選項
    echo 📥 下載地址: https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

:: 創建虛擬環境
echo 🔧 創建虛擬環境...
python -m venv ims_env

:: 啟動虛擬環境
echo 🔥 啟動虛擬環境...
call ims_env\Scripts\activate.bat

:: 升級 pip
echo ⬆️ 升級 pip...
python -m pip install --upgrade pip

:: 安裝依賴
echo 📦 安裝相依套件...
pip install reportlab

:: 檢查安裝是否成功
echo ✅ 檢查安裝...
python -c "import reportlab; print('ReportLab 版本:', reportlab.Version)"

echo.
echo 🎉 安裝完成！
echo.
echo 📖 使用方法：
echo 1. 啟動虛擬環境: ims_env\Scripts\activate.bat
echo 2. 執行程式: python pdf_generator_tkinter.py
echo 3. 退出虛擬環境: deactivate
echo.
echo ⚠️  請確保 ims_list.json 檔案存在於專案目錄中
pause 