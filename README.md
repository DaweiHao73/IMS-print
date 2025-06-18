# IMS Print - 調貨單 PDF 生成器

一個基於 Python Tkinter 的調貨單 PDF 生成器，支援從 IMS 商品清單自動查詢商品資訊並生成專業的調貨單文件。

![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter-green.svg)  
![ReportLab](https://img.shields.io/badge/PDF-ReportLab-red.svg)
![Cross Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)

## 📋 功能特色

### 🎯 核心功能

- **智能商品查詢**: 自動從 `ims_list.json` 查詢商品描述
- **專業 PDF 生成**: 生成格式化的調貨單 PDF 文件
- **跨平台支援**: Windows、macOS、Linux 全平台支援
- **中文字體支援**: 自動偵測並載入系統中文字體
- **直觀操作介面**: 現代化的 GUI 界面設計

### 🛠️ 進階功能

- **即時預覽**: 生成前可預覽調貨單內容
- **智能驗證**: 輸入資料驗證與錯誤提示
- **商品管理**: 支援新增、編輯、移除商品項目
- **自動儲存路徑**: 記住上次選擇的儲存位置
- **狀態回饋**: 即時顯示操作狀態與進度

## 🚀 快速開始

### 📋 系統需求

- Python 3.7 或更高版本（**必須包含 tkinter 模組**）
- 支援的作業系統：Windows、macOS、Linux

> **重要提醒**:
>
> - **Windows 用戶**: 安裝 Python 時請確保勾選「tcl/tk and IDLE」選項
> - **Linux 用戶**: 可能需要額外安裝 `python3-tk` 套件
> - **macOS 用戶**: 系統預設 Python 已包含 tkinter

### 📦 安裝步驟

1. **克隆專案**

   ```bash
   git clone https://github.com/your-username/IMS-print.git
   cd IMS-print
   ```

2. **自動安裝（推薦）**

   **macOS/Linux:**

   ```bash
   chmod +x setup.sh
   ./setup.sh
   ```

   **Windows:**

   ```batch
   setup.bat
   ```

   **手動安裝:**

   ```bash
   # 創建虛擬環境
   python3 -m venv ims_env

   # 啟動虛擬環境
   # macOS/Linux:
   source ims_env/bin/activate
   # Windows:
   ims_env\Scripts\activate.bat

   # 安裝相依套件
   pip install -r requirements.txt
   ```

3. **準備商品資料檔案**
   - 確保 `ims_list.json` 檔案位於專案根目錄
   - 檔案格式應符合以下結構：
   ```json
   [
     {
       "Item No": "商品編號",
       "Item Description": "商品描述"
     }
   ]
   ```

### 🏃‍♂️ 執行應用程式

#### 方法一：直接執行 Python 檔案

```bash
python pdf_generator_tkinter.py
```

#### 方法二：使用 main.py（功能更完整）

```bash
python main.py
```

#### 方法三：使用執行腳本（macOS/Linux）

```bash
chmod +x run_app.sh
./run_app.sh
```

## 📖 使用說明

### 1. 基本資訊填寫

在左側「基本資訊」區域填寫以下資料：

- **日期**: 調貨日期（預設為今日）
- **寄出店別**: 調出商品的店別
- **寄件人**: 調出商品的負責人
- **收件店別**: 接收商品的店別
- **收件人**: 接收商品的負責人
- **備註**: 額外備註資訊
- **儲存路徑**: 選擇 PDF 檔案儲存位置

### 2. 商品資訊輸入

在右側「商品輸入」區域：

- 輸入**商品編號**後按 Tab 或點擊「查詢」
- 系統會自動填入**商品描述**（從 ims_list.json 查詢）
- 輸入**數量**後點擊「加入」將商品加入清單

### 3. 商品清單管理

- **檢視**: 所有已加入的商品會顯示在清單中
- **編輯**: 雙擊商品項目可重新編輯
- **移除**: 選取商品後點擊「移除選取」
- **清空**: 點擊「清空全部」清空所有商品

### 4. PDF 生成

1. 點擊「預覽資料」檢查輸入內容
2. 點擊「產生 PDF」生成調貨單
3. 選擇是否開啟生成的 PDF 檔案
4. 選擇是否開啟儲存資料夾

## 🗂️ 檔案結構

```
IMS-print/
├── pdf_generator_tkinter.py    # 主要應用程式（簡化版）
├── main.py                     # 完整功能版本
├── ims_list.json              # 商品資料檔案
├── requirements.txt           # Python 依賴套件清單
├── setup.sh                   # macOS/Linux 自動安裝腳本
├── setup.bat                  # Windows 自動安裝腳本
├── run_app.sh                 # macOS/Linux 執行腳本
├── run_pdf_generator.ps1      # Windows PowerShell 腳本
└── README.md                  # 說明文件
```

## ⚙️ 配置說明

### 字體設定

程式會自動偵測並載入系統字體：

**Windows 系統**:

- 微軟正黑體 (`msjh.ttc`)
- 微軟雅黑 (`msyh.ttc`)
- 黑體 (`simhei.ttf`)
- 宋體 (`simsun.ttc`)

**macOS 系統**:

- 蘋方 (`PingFang.ttc`)
- Helvetica (`Helvetica.ttc`)
- 宋體 (`Songti.ttc`)

**Linux 系統**:

- Liberation Sans
- DejaVu Sans
- Noto Sans CJK

### IMS 資料格式

`ims_list.json` 檔案必須包含以下欄位：

```json
[
  {
    "Item No": "30495",
    "Item Description": "BROCHURE RACK F GRID CEILING 1-COMPARTMENT A4/LETTER H1700-1900MM WHI"
  }
]
```

## 🐛 故障排除

### 常見問題

**Q: 程式啟動時顯示「未找到商品資料檔案」**

- 確認 `ims_list.json` 檔案位於程式相同目錄
- 檢查檔案名稱拼寫是否正確
- 確認檔案格式為有效的 JSON

**Q: PDF 中中文字顯示為方框**

- 確認系統已安裝中文字體
- Windows: 確認微軟正黑體或其他中文字體已安裝
- macOS: 系統預設已包含中文字體
- Linux: 安裝 `fonts-noto-cjk` 套件

**Q: 無法生成 PDF 檔案**

- 檢查儲存路徑是否有寫入權限
- 確認 reportlab 套件已正確安裝
- 檢查磁碟空間是否足夠

**Q: 商品查詢功能無法使用**

- 確認商品編號輸入正確
- 檢查 `ims_list.json` 檔案中是否包含該商品
- 確認檔案編碼為 UTF-8

**Q: Windows 上出現「No module named '\_tkinter'」錯誤**

- **方法一 - 重新安裝 Python（推薦）:**

  1. 下載最新版 Python: https://www.python.org/downloads/windows/
  2. 安裝時確保勾選「Add Python to PATH」
  3. 安裝時確保勾選「tcl/tk and IDLE」選項
  4. 完整安裝後重新執行程式

- **方法二 - 透過 Microsoft Store 安裝:**

  ```batch
  # 開啟 Microsoft Store 搜尋並安裝 Python
  # 或使用 winget 指令
  winget install Python.Python.3.11
  ```

- **方法三 - 檢查現有安裝:**

  ```batch
  # 檢查 Python 安裝位置
  where python

  # 檢查是否包含 tkinter
  python -c "import tkinter; print('tkinter 可用')"
  ```

- **方法四 - 透過 conda 安裝 (如果使用 Anaconda):**
  ```batch
  conda install tk
  ```

## 🛡️ 系統需求

### 最低需求

- **Python**: 3.7+
- **記憶體**: 512 MB RAM
- **儲存空間**: 50 MB 可用空間
- **解析度**: 1024x768 或更高

### 建議配置

- **Python**: 3.9+
- **記憶體**: 2 GB RAM
- **儲存空間**: 100 MB 可用空間
- **解析度**: 1920x1080

## 📚 相依套件

| 套件名稱  | 版本需求 | 用途         |
| --------- | -------- | ------------ |
| reportlab | >= 3.5.0 | PDF 文件生成 |
| tkinter   | 內建     | GUI 界面框架 |

## 🔄 版本歷史

### v1.2.0 (最新)

- ✨ 新增資料預覽功能
- ✨ 改善 UI 佈局與使用體驗
- ✨ 加入狀態列顯示操作回饋
- 🐛 修復商品重複加入的問題
- 🐛 改善跨平台字體載入

### v1.1.0

- ✨ 支援商品項目編輯功能
- ✨ 加入表單驗證機制
- 🔧 優化 PDF 生成效能

### v1.0.0

- 🎉 初始版本發布
- ✨ 基本調貨單生成功能
- ✨ IMS 商品查詢整合

## 🤝 貢獻指南

歡迎提交 Issue 和 Pull Request！

### 開發環境設置

1. Fork 此專案
2. 創建功能分支：`git checkout -b feature/AmazingFeature`
3. 提交更改：`git commit -m 'Add some AmazingFeature'`
4. 推送分支：`git push origin feature/AmazingFeature`
5. 開啟 Pull Request

### 程式碼規範

- 遵循 PEP 8 程式碼風格
- 提供適當的註解和文件
- 確保跨平台相容性

## 📄 授權條款

此專案使用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案

## 📞 聯絡資訊

- **開發者**: Your Name
- **Email**: your.email@example.com
- **專案首頁**: https://github.com/your-username/IMS-print

## 🙏 致謝

感謝以下開源專案：

- [ReportLab](https://www.reportlab.com/) - PDF 生成引擎
- [Tkinter](https://docs.python.org/3/library/tkinter.html) - GUI 框架

---

**⭐ 如果這個專案對您有幫助，請給一個 Star！**
