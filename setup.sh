#!/bin/bash

# IMS Print 安裝腳本
echo "🚀 開始安裝 IMS Print 調貨單生成器..."

# 檢查 Python 版本
python_version=$(python3 --version 2>&1 | awk '{print $2}' | awk -F. '{print $1"."$2}')
echo "📋 檢測到 Python 版本: $python_version"

if [[ $(echo "$python_version >= 3.7" | bc -l) -eq 0 ]]; then
    echo "❌ 需要 Python 3.7 或更高版本"
    exit 1
fi

# 檢查 tkinter 是否可用
echo "🔍 檢查 tkinter 模組..."
if python3 -c "import tkinter" 2>/dev/null; then
    echo "✅ tkinter 模組可用"
else
    echo "❌ tkinter 模組不可用"
    echo "💡 Linux 用戶請執行: sudo apt-get install python3-tk"
    echo "💡 macOS 用戶請重新安裝 Python 或使用 Homebrew: brew install python-tk"
    exit 1
fi

# 創建虛擬環境
echo "🔧 創建虛擬環境..."
python3 -m venv ims_env

# 啟動虛擬環境
echo "🔥 啟動虛擬環境..."
source ims_env/bin/activate

# 升級 pip
echo "⬆️ 升級 pip..."
python -m pip install --upgrade pip

# 安裝依賴
echo "📦 安裝相依套件..."
pip install reportlab

# 檢查安裝是否成功
echo "✅ 檢查安裝..."
python -c "import reportlab; print('ReportLab 版本:', reportlab.Version)"

echo "🎉 安裝完成！"
echo ""
echo "📖 使用方法："
echo "1. 啟動虛擬環境: source ims_env/bin/activate"
echo "2. 執行程式: python pdf_generator_tkinter.py"
echo "3. 退出虛擬環境: deactivate"
echo ""
echo "⚠️  請確保 ims_list.json 檔案存在於專案目錄中" 