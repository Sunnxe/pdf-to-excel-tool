# 📄→📊 PDF轉Excel工具

一個基於Flask的PDF工單轉Excel工具，支援拖拽上傳並自動分類材料代碼。

## 🚀 功能特色

- **📄 PDF解析** - 使用pdfplumber精確解析工單PDF
- **🎯 智能分類** - 自動分類H系列原料和I系列鐵材代碼
- **📊 Excel輸出** - 生成分類完整的Excel檔案
- **🖱️ 拖拽上傳** - 直接拖拽PDF檔案到網頁
- **⚡ 即時處理** - 快速轉換並下載結果

## 📦 部署到Railway

### 方法1: 直接從GitHub部署

1. **Fork此倉庫**到你的GitHub帳號
2. 前往 [railway.app](https://railway.app)
3. 點擊 "New Project" → "Deploy from GitHub repo"
4. 選擇你fork的倉庫
5. Railway會自動部署，獲得永久網址

### 方法2: 手動上傳

1. 下載此倉庫的所有檔案
2. 創建新的GitHub倉庫
3. 上傳所有檔案到倉庫
4. 連接Railway進行部署

## 📁 檔案結構

```
📁 專案根目錄
├── 📄 app.py              # Flask主應用程式
├── 📄 requirements.txt    # Python依賴套件
├── 📄 Procfile           # 啟動配置
├── 📄 railway.json       # Railway部署配置
└── 📁 final/             # PDF抽取器模組
    ├── 📄 __init__.py
    └── 📄 pdf_extractor.py
```

## 🛠️ 本地開發

```bash
# 安裝依賴
pip install -r requirements.txt

# 啟動應用
python app.py

# 訪問 http://localhost:5000
```

## 📊 支援的PDF格式

- ✅ 工單明細表PDF
- ✅ 包含PD開頭工單號的PDF
- ✅ 電腦生成的結構化PDF
- ✅ 包含H系列材料代碼的PDF
- ✅ 包含I系列鐵材代碼的PDF

## 💡 使用方式

1. **訪問網站** - 打開部署後的Railway網址
2. **上傳PDF** - 拖拽或選擇PDF檔案
3. **等待處理** - 系統自動解析PDF內容
4. **下載Excel** - 處理完成後自動下載

## 🎯 Excel輸出格式

| 欄位名稱 | 說明 | 範例 |
|---------|------|------|
| H系列代碼 | H開頭原料代碼 | HS-C9-45-02-B |
| 原料公斤數 | 純數字重量 | 15.8 |
| I系列代碼 | I開頭鐵材代碼 | IAAD003404800542z |
| 鐵材隻數 | 純數字數量 | 4.0 |
| 其他材料 | 其他類型材料 | g(8.0) |

## 💰 費用

- **Railway免費額度**: 每月$5
- **預估使用量**: 每次轉換約$0.01
- **免費處理數**: 約500次/月

## 🔧 技術架構

- **後端**: Flask + Python
- **PDF解析**: pdfplumber
- **Excel生成**: pandas + openpyxl
- **前端**: HTML + CSS + JavaScript
- **部署**: Railway

## 📞 支援

如遇到問題：
1. 檢查PDF是否為工單格式
2. 確認包含PD開頭的工單號
3. 檢查PDF是否為文字型（非掃描圖片）

---

**立即部署到Railway，開始使用！** 🚀