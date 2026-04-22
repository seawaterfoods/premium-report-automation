# 系統架構文件

> 本文件包含以 Mermaid 撰寫的架構圖、時序圖、邏輯圖。
> 可在 GitHub、VS Code (Mermaid Preview 插件)、或任何 Mermaid 渲染器中檢視。

---

## 一、系統架構圖 (Architecture)

```mermaid
graph TB
    subgraph Input["📂 來源資料夾 (import/)"]
        SRC1["01_11501_保險收入統計表.xlsx"]
        SRC2["02_11501_保險收入統計表.xlsx"]
        SRCN["...N 家公司 × M 月份"]
    end

    subgraph App["🖥️ Spring Boot Application"]
        CONFIG["AppConfig<br/>application.yml"]
        
        subgraph Reader["📖 Reader Layer"]
            SCANNER["FileScanner<br/>掃描 & 探索檔案"]
            VALIDATOR["FileValidator<br/>結構 & 命名驗證"]
            PARSER["ExcelSourceReader<br/>解析 Excel 來源"]
        end
        
        subgraph Model["📦 Domain Model"]
            COMPANY["CompanyData"]
            INSURANCE["InsuranceRecord"]
            CATEGORY["CategoryMapping<br/>歸屬對照 (從YAML載入)"]
        end
        
        subgraph Calculator["🧮 Calculator Layer"]
            MONTHLY["MonthlyCalculator<br/>單月彙整"]
            CUMULATIVE["CumulativeCalculator<br/>累計彙整"]
            CATCALC["CategoryCalculator<br/>九大類加總"]
            COMPARISON["ComparisonCalculator<br/>同期比較 & 成長率"]
        end
        
        subgraph Writer["📝 Writer Layer"]
            W1["PremiumReportWriter<br/>保費統計表 5 頁簽"]
            W2["ComparisonReportWriter<br/>同期比較表 2 頁簽"]
            FORMATTER["ExcelStyleHelper<br/>格式復刻"]
        end
        
        SERVICE["ReportGenerationService<br/>ETL 流程協調器"]
    end

    subgraph Output["📂 輸出資料夾 (output/{年月}/)"]
        OUT1["115年產險業務(簽單)<br/>保費統計表.xlsx"]
        OUT2["11503vs11403同期<br/>比較分析表.xlsx"]
    end

    LOG["output/report.log<br/>(每次覆蓋)"]

    SRC1 --> SCANNER
    SRC2 --> SCANNER
    SRCN --> SCANNER
    
    CONFIG --> SERVICE
    SERVICE --> SCANNER
    SCANNER --> VALIDATOR
    VALIDATOR --> PARSER
    PARSER --> COMPANY
    PARSER --> INSURANCE
    PARSER --> CATEGORY
    
    SERVICE --> MONTHLY
    SERVICE --> CUMULATIVE
    SERVICE --> CATCALC
    SERVICE --> COMPARISON
    
    MONTHLY --> W1
    CUMULATIVE --> W1
    CATCALC --> W1
    COMPARISON --> W2
    
    W1 --> FORMATTER
    W2 --> FORMATTER
    
    W1 --> OUT1
    W2 --> OUT2
    SERVICE --> LOG
```

---

## 二、ETL 處理時序圖 (Sequence)

```mermaid
sequenceDiagram
    participant BAT as run.bat
    participant APP as Application
    participant SVC as ReportGenerationService
    participant SCN as FileScanner
    participant VAL as FileValidator
    participant RDR as ExcelSourceReader
    participant CAL as Calculators
    participant WRT as Writers
    participant FS as FileSystem

    BAT->>APP: java -jar xxx.jar
    APP->>SVC: execute()
    
    Note over SVC: Step 1: 讀取設定
    SVC->>SVC: 讀取 application.yml
    
    Note over SVC,FS: Step 1: 掃描來源資料夾
    SVC->>SCN: scan(importDir, processYear)
    SCN->>FS: 列舉 import/{年份}/{月份}/*.xlsx
    FS-->>SCN: 檔案清單
    SCN->>SCN: 依檔名區分今年/去年
    SCN->>SCN: 偵測最新月份
    SCN-->>SVC: SourceFileSet (今年+去年)
    
    Note over SVC,RDR: Step 2: 來源檔案內容檢核
    loop 每份來源檔
        SVC->>RDR: validateContent(file)
        RDR->>RDR: 檢查公式 (C2, B4, C6:C38)
        RDR->>RDR: 檢查小數 (C6:C38)
        RDR->>RDR: 比對 C2 年月 vs 檔名
        RDR->>RDR: 比對 B4 代號 vs 檔名
        RDR-->>SVC: 錯誤清單
    end
    alt 有檢核異常
        SVC->>SVC: 全部列出至 report.log
        SVC-->>APP: 中止 (不產生報表)
    end

    Note over SVC,RDR: Step 3: 讀取資料
    loop 每份來源檔
        SVC->>VAL: validate(file)
        VAL->>VAL: 驗證檔名/結構/代號
        VAL-->>SVC: ValidationResult
        SVC->>RDR: read(file)
        RDR->>RDR: 解析 Sheet1 (金額)
        RDR->>RDR: 解析 Sheet2 (歸屬)
        RDR->>RDR: 正規化險種代號 → 文字
        RDR-->>SVC: CompanyData
    end
    
    Note over SVC,CAL: Step 4: 計算
    SVC->>CAL: calculateMonthly(allData)
    CAL-->>SVC: 單月明細 (無資料補0)
    SVC->>CAL: calculateCumulative(allData)
    CAL-->>SVC: 累計明細 (至最新月份)
    SVC->>CAL: calculateCategories(allData)
    CAL-->>SVC: 九大類彙總
    SVC->>CAL: calculateComparison(thisYear, lastYear)
    Note right of CAL: 去年=0 或 <0 → 分母=1, 記錄警告
    CAL-->>SVC: 同期比較 + 成長率
    
    Note over SVC,FS: Step 5: 輸出 Excel
    SVC->>WRT: writePremiumReport(data)
    WRT->>WRT: Sheet1: {YYY}單
    WRT->>WRT: Sheet2: {YYY}單累
    WRT->>WRT: Sheet3: {YYY}總
    WRT->>WRT: Sheet4: {YYY}總累
    WRT->>WRT: Sheet5: 歸屬
    WRT->>WRT: 套用格式 (合併儲存格/框線/粗體)
    WRT->>FS: 儲存至 output/{年月}/ .xlsx
    
    SVC->>WRT: writeComparisonReport(data)
    WRT->>WRT: Sheet1: 比較增減率
    WRT->>WRT: Sheet2: 增減原因 (E欄留空)
    WRT->>FS: 儲存至 output/{年月}/ .xlsx
    
    Note over SVC: Step 6: 完成
    SVC->>SVC: 輸出執行摘要
    SVC-->>APP: done
    APP-->>BAT: exit 0
```

---

## 三、成長率計算邏輯圖 (Logic)

```mermaid
flowchart TD
    START([計算成長率]) --> CHECK_DENOM{去年數值}
    
    CHECK_DENOM -->|去年 > 0| NORMAL[分母 = 去年值]
    CHECK_DENOM -->|去年 = 0| ZERO[分母 = 1<br/>⚠️ 記錄警告：缺資料]
    CHECK_DENOM -->|去年 < 0| NEGATIVE[分母 = 1<br/>⚠️ 記錄警告：負數資料]
    
    NORMAL --> CALC["成長率 = (今年 / 分母) - 1"]
    ZERO --> CALC
    NEGATIVE --> CALC
    
    CALC --> ROUND["四捨五入至小數第 2 位"]
    ROUND --> DONE([回傳成長率 %])
```

---

## 四、資料流程圖 (Data Flow)

```mermaid
flowchart LR
    subgraph Source["來源 (N份)"]
        S["保險收入統計表.xlsx<br/>33 險種 × 1 公司 × 1 月"]
    end
    
    subgraph Transform["轉換"]
        direction TB
        T1["合併多公司多月份"]
        T2["正規化險種代號"]
        T3["歸屬分類 → 九大類"]
        T4["單月/累計計算"]
        T5["同期比較"]
        T1 --> T2 --> T3 --> T4 --> T5
    end
    
    subgraph Output1["輸出1: 保費統計表 (output/{年月}/)"]
        O1A["{YYY}單<br/>N公司×12月×33險種"]
        O1B["{YYY}單累<br/>N公司×累計×33險種"]
        O1C["{YYY}總<br/>N公司×12月×16類"]
        O1D["{YYY}總累<br/>N公司×累計×16類"]
        O1E["歸屬<br/>分類對照表"]
    end
    
    subgraph Output2["輸出2: 同期比較 (output/{年月}/)"]
        O2A["比較增減率<br/>15類×單月+累計"]
        O2B["增減原因<br/>15類×成長率+空白"]
    end
    
    S --> T1
    T5 --> O1A
    T5 --> O1B
    T5 --> O1C
    T5 --> O1D
    T5 --> O1E
    T5 --> O2A
    T5 --> O2B
```

---

## 五、套件結構圖 (Package)

```mermaid
graph TB
    subgraph com.insurance.report
        APP["PremiumReportAutomationApplication<br/>🚀 進入點 + CommandLineRunner"]
        
        subgraph config
            AC["AppConfig<br/>📋 application.yml 映射"]
        end
        
        subgraph model
            M1["CompanyData<br/>公司 + 月份 + 金額"]
            M2["InsuranceCode<br/>險種代號枚舉"]
            M3["CategoryMapping<br/>歸屬對照 (從YAML載入)"]
            M4["MonthlyReport<br/>單月/累計資料"]
            M5["ComparisonResult<br/>同期比較結果"]
        end
        
        subgraph reader
            R1["FileScanner<br/>掃描 import/"]
            R2["FileValidator<br/>驗證結構"]
            R3["ExcelSourceReader<br/>解析 Excel"]
        end
        
        subgraph calculator
            C1["MonthlyCalculator"]
            C2["CumulativeCalculator"]
            C3["CategoryCalculator"]
            C4["ComparisonCalculator"]
        end
        
        subgraph writer
            W1["PremiumReportWriter<br/>報表一 (5 sheets)"]
            W2["ComparisonReportWriter<br/>報表二 (2 sheets)"]
            W3["ExcelStyleHelper<br/>格式工具"]
        end
        
        subgraph service
            SVC["ReportGenerationService<br/>ETL 主流程"]
        end
        
        subgraph util
            U1["InsuranceCodeUtil<br/>代號正規化"]
            U2["GrowthRateUtil<br/>成長率計算"]
        end
    end
    
    APP --> SVC
    SVC --> AC
    SVC --> R1
    SVC --> C1
    SVC --> W1
    SVC --> W2
```
