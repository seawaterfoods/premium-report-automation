package com.insurance.report.service;

import com.insurance.report.calculator.*;
import com.insurance.report.config.AppConfig;
import com.insurance.report.model.*;
import com.insurance.report.reader.*;
import com.insurance.report.writer.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

/**
 * 報表產生主服務 — 協調整個 ETL 流程
 */
@Service
public class ReportGenerationService {

    private static final Logger log = LoggerFactory.getLogger(ReportGenerationService.class);

    private final AppConfig config;
    private final FileScanner fileScanner;
    private final FileValidator fileValidator;
    private final ExcelSourceReader excelReader;
    private final MonthlyCalculator monthlyCalculator;
    private final CumulativeCalculator cumulativeCalculator;
    private final CategoryCalculator categoryCalculator;
    private final ComparisonCalculator comparisonCalculator;
    private final RankingCalculator rankingCalculator;
    private final PremiumReportWriter premiumWriter;
    private final ComparisonReportWriter comparisonWriter;

    public ReportGenerationService(
            AppConfig config,
            FileScanner fileScanner,
            FileValidator fileValidator,
            ExcelSourceReader excelReader,
            MonthlyCalculator monthlyCalculator,
            CumulativeCalculator cumulativeCalculator,
            CategoryCalculator categoryCalculator,
            ComparisonCalculator comparisonCalculator,
            RankingCalculator rankingCalculator,
            PremiumReportWriter premiumWriter,
            ComparisonReportWriter comparisonWriter) {
        this.config = config;
        this.fileScanner = fileScanner;
        this.fileValidator = fileValidator;
        this.excelReader = excelReader;
        this.monthlyCalculator = monthlyCalculator;
        this.cumulativeCalculator = cumulativeCalculator;
        this.categoryCalculator = categoryCalculator;
        this.comparisonCalculator = comparisonCalculator;
        this.rankingCalculator = rankingCalculator;
        this.premiumWriter = premiumWriter;
        this.comparisonWriter = comparisonWriter;
    }

    public void execute() throws IOException {
        int year = config.getProcessYear();
        int priorYear = year - 1;

        // Step 1: 掃描來源資料夾
        log.info("Step 1: 掃描來源資料夾");
        SourceFileSet fileSet = fileScanner.scan();
        if (fileSet.getCurrentYearFiles().isEmpty()) {
            throw new IOException("找不到任何今年(" + year + ")的來源檔案");
        }
        int latestMonth = fileSet.getLatestMonth();

        // Step 2: 驗證 & 讀取
        log.info("Step 2: 驗證 & 讀取 (今年 {} 筆, 去年 {} 筆)",
                countFiles(fileSet.getCurrentYearFiles()),
                countFiles(fileSet.getPriorYearFiles()));

        List<CompanyMonthData> currentYearData = readAllFiles(fileSet.getCurrentYearFiles());
        List<CompanyMonthData> priorYearData = readAllFiles(fileSet.getPriorYearFiles());

        if (priorYearData.isEmpty()) {
            log.warn("找不到去年({})資料，去年數值將預設為 0", priorYear);
        }

        // 收集公司資訊
        Map<String, String> companyNames = new LinkedHashMap<>();
        currentYearData.forEach(d -> companyNames.putIfAbsent(d.getCompanyCode(), d.getCompanyName()));
        List<String> companyList = sortCompanyList(new ArrayList<>(companyNames.keySet()));

        log.info("公司清單 ({}家): {}", companyList.size(), companyList);

        // Step 3: 計算
        log.info("Step 3: 計算 (最新月份: {}月)", latestMonth);

        // 3a: 單月彙整
        Map<Integer, List<CompanyMonthData>> monthlyData =
                monthlyCalculator.calculate(currentYearData, companyList);
        Map<Integer, CompanyMonthData> monthlySubtotals = new LinkedHashMap<>();
        for (var entry : monthlyData.entrySet()) {
            monthlySubtotals.put(entry.getKey(),
                    monthlyCalculator.calculateSubtotal(entry.getValue(), year, entry.getKey()));
        }

        // 3b: 累計彙整
        Map<Integer, List<CompanyMonthData>> cumulativeData =
                cumulativeCalculator.calculate(monthlyData, companyList, latestMonth);
        Map<Integer, CompanyMonthData> cumulativeSubtotals = new LinkedHashMap<>();
        for (var entry : cumulativeData.entrySet()) {
            cumulativeSubtotals.put(entry.getKey(),
                    cumulativeCalculator.calculateSubtotal(entry.getValue(), year, entry.getKey()));
        }

        // 3c: 去年同期
        Map<Integer, List<CompanyMonthData>> priorMonthlyData =
                monthlyCalculator.calculate(priorYearData, sortCompanyList(getCompanyCodes(priorYearData)));
        Map<Integer, CompanyMonthData> priorMonthlySubtotals = new LinkedHashMap<>();
        for (var entry : priorMonthlyData.entrySet()) {
            priorMonthlySubtotals.put(entry.getKey(),
                    monthlyCalculator.calculateSubtotal(entry.getValue(), priorYear, entry.getKey()));
        }

        Map<Integer, List<CompanyMonthData>> priorCumulativeData =
                cumulativeCalculator.calculate(priorMonthlyData,
                        sortCompanyList(getCompanyCodes(priorYearData)), latestMonth);
        Map<Integer, CompanyMonthData> priorCumulativeSubtotals = new LinkedHashMap<>();
        for (var entry : priorCumulativeData.entrySet()) {
            priorCumulativeSubtotals.put(entry.getKey(),
                    cumulativeCalculator.calculateSubtotal(entry.getValue(), priorYear, entry.getKey()));
        }

        // 3d: 同期比較 (比較表用)
        Map<String, Long> currentMonthCatTotals = calculateCategorySubtotals(monthlySubtotals.get(latestMonth));
        Map<String, Long> priorMonthCatTotals = calculateCategorySubtotals(priorMonthlySubtotals.get(latestMonth));
        Map<String, Long> currentCumCatTotals = calculateCategorySubtotals(cumulativeSubtotals.get(latestMonth));
        Map<String, Long> priorCumCatTotals = calculateCategorySubtotals(priorCumulativeSubtotals.get(latestMonth));

        List<String> comparisonCategoryNames = new ArrayList<>();
        for (CategoryMapping.SubCategory sub : CategoryMapping.SUB_CATEGORIES) {
            comparisonCategoryNames.add(sub.getName());
        }

        ComparisonCalculator.ComparisonRow monthlyComparison = comparisonCalculator.calculate(
                currentMonthCatTotals, priorMonthCatTotals, comparisonCategoryNames,
                latestMonth + "月");
        ComparisonCalculator.ComparisonRow cumulativeComparison = comparisonCalculator.calculate(
                currentCumCatTotals, priorCumCatTotals, comparisonCategoryNames,
                "1-" + latestMonth + "月累計");

        // 3e: 排名 (年度累計至最新月份)
        Map<CompanyInfo, Long> rankingData = new LinkedHashMap<>();
        List<CompanyMonthData> latestCumulative = cumulativeData.get(latestMonth);
        if (latestCumulative != null) {
            for (CompanyMonthData d : latestCumulative) {
                rankingData.put(new CompanyInfo(d.getCompanyCode(), d.getCompanyName()), d.getTotal());
            }
        }
        List<RankingCalculator.RankEntry> ranking = rankingCalculator.calculate(rankingData);

        // Step 4: 輸出 Excel
        log.info("Step 4: 輸出 Excel");
        Path outputDir = Paths.get(config.getOutputDir());

        // 報表一：保費統計表
        String report1Name = String.format("%d年產險業務(簽單)保費統計表.xlsx", year);
        premiumWriter.write(year, latestMonth,
                monthlyData, cumulativeData,
                monthlySubtotals, cumulativeSubtotals,
                priorMonthlySubtotals, priorCumulativeSubtotals,
                ranking, outputDir.resolve(report1Name));

        // 報表二：同期比較分析表
        int yearMonth = year * 100 + latestMonth;
        int priorYearMonth = priorYear * 100 + latestMonth;
        String report2Name = String.format("%05dvs%05d同期比較分析表.xlsx", yearMonth, priorYearMonth);
        comparisonWriter.write(year, latestMonth,
                monthlyComparison, cumulativeComparison,
                outputDir.resolve(report2Name));

        // Step 5: 執行摘要
        log.info("========== 執行摘要 ==========");
        log.info("處理年度: {} (去年: {})", year, priorYear);
        log.info("公司數: {} 家", companyList.size());
        log.info("月份: 1~{} (今年 {} 筆, 去年 {} 筆)",
                latestMonth, currentYearData.size(), priorYearData.size());
        log.info("輸出: {}", outputDir.toAbsolutePath());
        log.info("  - {}", report1Name);
        log.info("  - {}", report2Name);
    }

    // --- 內部方法 ---

    private List<CompanyMonthData> readAllFiles(Map<Integer, List<SourceFileInfo>> filesByMonth) {
        List<CompanyMonthData> result = new ArrayList<>();
        for (var entry : filesByMonth.entrySet()) {
            for (SourceFileInfo fileInfo : entry.getValue()) {
                List<String> errors = fileValidator.validate(fileInfo);
                if (!errors.isEmpty()) {
                    log.error("檔案驗證失敗: {}", errors);
                    continue;
                }
                try {
                    result.add(excelReader.read(fileInfo));
                } catch (Exception e) {
                    log.error("讀取檔案失敗: {} - {}", fileInfo.getFilePath(), e.getMessage());
                }
            }
        }
        return result;
    }

    private Map<String, Long> calculateCategorySubtotals(CompanyMonthData subtotal) {
        if (subtotal == null) return Collections.emptyMap();
        return categoryCalculator.calculateSubCategories(subtotal);
    }

    private List<String> getCompanyCodes(List<CompanyMonthData> data) {
        Set<String> codes = new TreeSet<>();
        data.forEach(d -> codes.add(d.getCompanyCode()));
        return new ArrayList<>(codes);
    }

    private List<String> sortCompanyList(List<String> codes) {
        switch (config.getCompanyOrder()) {
            case BY_CODE_DESC:
                codes.sort(Comparator.reverseOrder());
                break;
            case BY_NAME:
            case BY_CODE_ASC:
            default:
                Collections.sort(codes);
                break;
        }
        return codes;
    }

    private int countFiles(Map<Integer, List<SourceFileInfo>> filesByMonth) {
        return filesByMonth.values().stream().mapToInt(List::size).sum();
    }
}
