# 學生積分管理系統 - 專案架構說明

## 專案概述
本系統用於自動化管理學生出缺席打卡紀錄，並將其轉化為積分銀行中的獎勵點數。

## 目錄結構
```text
學生積分系統/
├── README.txt               # 本說明文件
├── main.py                  # 系統整合入口，執行爬蟲與點數計算
├── scrapers/                # 爬蟲模組
│   └── attendance_scraper.py # 自動登入校務系統抓取到離校時間
├── gas/                     # Google Apps Script 雲端邏輯
│   ├── Code.js              # 積分查詢與 API 處理邏輯
│   └── Index.html           # 學生查詢介面前端
├── data/                    # 資料儲存區
│   ├── 學生出缺席_總表.xlsx     # 爬蟲抓取的原始打卡資料 (多分頁)
│   └── 106班積分銀行...csv   # 積分銀行的學生基準資料
└── utils/                   # 工具程式
    └── point_calculator.py  # 負責解析 Excel 並計算累計點數
```

## 運作流程
1. **資料抓取**：執行 `scrapers/attendance_scraper.py`，從校務系統抓取資料並存入 `data/學生出缺席_總表.xlsx`。
2. **點數計算**：`utils/point_calculator.py` 讀取 Excel，只要學生在任一天有「到校時間」紀錄，即判定為打卡成功，加 1 點。
3. **整合同步**：透過 `main.py` 串接上述流程，並在未來透過 Google Sheets API 或 Web App 同步至雲端。

## 開發者備註
- 爬蟲使用 `DrissionPage` 庫。
- 點數計算邏輯目前定義為：有到校時間 = +1 點。
- 建議使用 `clasp` 管理 Google Apps Script 程式碼。
