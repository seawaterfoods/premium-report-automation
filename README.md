# 產險業務保費統計自動化系統

> 將多家保險公司的保險收入統計表自動彙整，產出保費統計表及同期比較分析表。

## 技術棧

- Java 17
- Spring Boot 3.5.0
- Apache POI 5.3.0 (Excel 處理)
- Maven (含 wrapper，免安裝)

## 快速開始

### 1. 編譯

```cmd
build.bat
```

### 2. 準備來源資料

將來源 Excel 放入 `import/` 資料夾，**今年和去年的資料都要放**：

```
import/
├── 115/                         ← 今年
│   ├── 01/
│   │   ├── 01_11501_保險收入統計表.xlsx
│   │   ├── 02_11501_保險收入統計表.xlsx
│   │   └── ...
│   ├── 02/
│   └── 03/
└── 114/                         ← 去年 (供同期比較用)
    ├── 01/
    ├── 02/
    └── 03/
```

檔名格式：`{公司代號}_{民國年月}_保險收入統計表.xlsx`

### 3. 確認設定

編輯 `config/application.yml`：

```yaml
app:
  process-year: 115              # 處理年份
  columns:
    hidden-codes: ["9900"]       # 隱藏的險種代號
  company-order: BY_CODE_ASC     # 公司排序
```

### 4. 執行

```cmd
run.bat
```

### 5. 輸出

結果產生於 `output/{年月}/` 資料夾，例如 `output/11503/`：
- `115年產險業務(簽單)保費統計表.xlsx` — 5 個頁簽
- `11503vs11403同期比較分析表.xlsx` — 2 個頁簽

執行紀錄：`output/report.log`（每次覆蓋）

## 設定

編輯 `config/application.yml` 調整：
- 來源/輸出路徑
- 處理年份
- 險種顯示/隱藏 (`hidden-codes`)
- 公司排序方式

詳見 [操作手冊](docs/user-guide.md)。

## 專案結構

```
premium-report-automation/
├── docs/                    ← 需求規格書 & 架構圖 & 操作手冊
├── config/                  ← 外部設定檔
├── import/                  ← 來源資料夾
├── output/                  ← 輸出資料夾 (按年月分子資料夾)
├── src/main/java/com/insurance/report/
│   ├── config/              ← 設定類別
│   ├── model/               ← 領域模型
│   ├── reader/              ← 來源檔案解析
│   ├── calculator/          ← ETL 計算引擎
│   ├── writer/              ← Excel 輸出
│   ├── service/             ← 流程協調
│   └── util/                ← 工具類別
├── run.bat                  ← 執行腳本
└── build.bat                ← 編譯腳本
```

## 文件

- [操作手冊](docs/user-guide.md) — 完整操作流程、設定說明、新增險種教學
- [需求規格書](docs/requirements-spec.md) — v1.0 (全部需求已確認)
- [架構圖](docs/architecture.md) — Mermaid 格式 (架構/時序/邏輯/資料流/套件圖)
