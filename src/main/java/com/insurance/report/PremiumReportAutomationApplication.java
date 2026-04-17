package com.insurance.report;

import com.insurance.report.service.ReportGenerationService;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class PremiumReportAutomationApplication implements CommandLineRunner {

	private static final Logger log = LoggerFactory.getLogger(PremiumReportAutomationApplication.class);

	private final ReportGenerationService reportGenerationService;

	public PremiumReportAutomationApplication(ReportGenerationService reportGenerationService) {
		this.reportGenerationService = reportGenerationService;
	}

	public static void main(String[] args) {
		SpringApplication.run(PremiumReportAutomationApplication.class, args);
	}

	@Override
	public void run(String... args) {
		log.info("========== 產險業務保費統計自動化系統 啟動 ==========");
		try {
			reportGenerationService.execute();
			log.info("========== 執行完成 ==========");
		} catch (Exception e) {
			log.error("執行失敗: {}", e.getMessage(), e);
			System.exit(1);
		}
	}
}
