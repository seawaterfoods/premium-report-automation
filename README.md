# 產險業務保費統計自動化系統

> 將多家保險公司的保險收入統計表自動彙整，產出保費統計表及同期比較分析表。

## 技術棧

- Java 17
- Spring Boot 3.5.0
- Apache POI 5.3.0 (Excel 處理)
- Maven (含 wrapper，免安裝)

## 快速開始

### 1. 編譯

```bash
# Windows
build.bat

# 或手動
mvnw.cmd clean package -DskipTests
```

### 2. 準備來源資料

將來源 Excel 放入 `import/` 資料夾：
```
import/
└── 115/
    ├── 01/
    │   ├── 01_11501_保險收入統計表.xlsx
    │   ├── 02_11501_保險收入統計表.xlsx
    │   └── ...
    ├── 02/
    └── 03/
```

檔名格式：`{公司代號}_{民國年月}_保險收入統計表.xlsx`

### 3. 執行

```bash
# 雙擊
run.bat

# 或手動
java -jar target/premium-report-automation-0.0.1-SNAPSHOT.jar
```

### 4. 輸出

結果產生於 `output/` 資料夾：
- `{YYY}年產險業務(簽單)保費統計表.xlsx` — 6 個頁簽
- `{YYMM}vs{YYMM-100}同期比較分析表.xlsx` — 2 個頁簽

## 設定

編輯 `config/application.yml` 調整：
- 來源/輸出路徑
- 處理年份
- 國外分進(9900) 是否納入
- 險種顯示/隱藏
- 公司排序方式

## 專案結構

```
premium-report-automation/
├── docs/                    ← 需求規格書 & 架構圖
├── config/                  ← 外部設定檔
├── import/                  ← 來源資料夾
├── output/                  ← 輸出資料夾
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

- [需求規格書](docs/requirements-spec.md) — v1.0 (全部需求已確認)
- [架構圖](docs/architecture.md) — Mermaid 格式 (架構/時序/邏輯/資料流/套件圖)
