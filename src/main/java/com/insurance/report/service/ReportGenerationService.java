package com.insurance.report.service;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

/**
 * 報表產生主服務 — 協調整個 ETL 流程
 */
@Service
public class ReportGenerationService {

    private static final Logger log = LoggerFactory.getLogger(ReportGenerationService.class);

    public void execute() {
        log.info("Step 1: 讀取設定");
        log.info("Step 2: 掃描來源資料夾");
        log.info("Step 3: 驗證來源檔案");
        log.info("Step 4: 讀取 & 解析");
        log.info("Step 5: 計算");
        log.info("Step 6: 輸出 Excel");
        log.info("Step 7: 完成");
        // TODO: 各階段實作將在後續 Stage 填入
    }
}
