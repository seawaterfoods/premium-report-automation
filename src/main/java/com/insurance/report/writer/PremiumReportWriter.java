package com.insurance.report.writer;

import com.insurance.report.calculator.*;
import com.insurance.report.config.AppConfig;
import com.insurance.report.model.*;
import com.insurance.report.model.CategoryMapping.SubCategory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

import static com.insurance.report.writer.ExcelStyleHelper.*;

/**
 * 保費統計表 Excel 輸出 (6 個頁簽)
 */
@Component
public class PremiumReportWriter {

    private static final Logger log = LoggerFactory.getLogger(PremiumReportWriter.class);

    private final AppConfig config;
    private final CategoryCalculator categoryCalculator;

    public PremiumReportWriter(AppConfig config, CategoryCalculator categoryCalculator) {
        this.config = config;
        this.categoryCalculator = categoryCalculator;
    }

    /**
     * 產生保費統計表
     */
    public void write(
            int year, int latestMonth,
            Map<Integer, List<CompanyMonthData>> monthlyData,
            Map<Integer, List<CompanyMonthData>> cumulativeData,
            Map<Integer, CompanyMonthData> monthlySubtotals,
            Map<Integer, CompanyMonthData> cumulativeSubtotals,
            Map<Integer, CompanyMonthData> priorMonthlySubtotals,
            Map<Integer, CompanyMonthData> priorCumulativeSubtotals,
            List<RankingCalculator.RankEntry> ranking,
            Path outputPath) throws IOException {

        log.info("產生保費統計表: {}", outputPath.getFileName());

        try (Workbook wb = new XSSFWorkbook()) {
            ExcelStyleHelper styles = new ExcelStyleHelper(wb);
            List<String> companyList = getCompanyList(monthlyData);

            // Sheet 1: {YYY}單
            writeDetailSheet(wb, styles, year + "單", year, monthlyData, monthlySubtotals,
                    companyList, false, 12);

            // Sheet 2: {YYY}單累
            writeDetailSheet(wb, styles, year + "單累", year, cumulativeData, cumulativeSubtotals,
                    companyList, true, latestMonth);

            // Sheet 3: {YYY}總
            writeSummarySheet(wb, styles, year + "總", year, latestMonth,
                    monthlyData, monthlySubtotals,
                    priorMonthlySubtotals,
                    companyList, false, 12);

            // Sheet 4: {YYY}總累
            writeSummarySheet(wb, styles, year + "總累", year, latestMonth,
                    cumulativeData, cumulativeSubtotals,
                    priorCumulativeSubtotals,
                    companyList, true, latestMonth);

            // Sheet 5: 歸屬
            writeGuishuSheet(wb, styles);

            // Sheet 6: 排名
            writeRankingSheet(wb, styles, year, latestMonth, ranking);

            // 儲存
            Files.createDirectories(outputPath.getParent());
            try (OutputStream os = Files.newOutputStream(outputPath)) {
                wb.write(os);
            }
        }

        log.info("保費統計表已產生: {}", outputPath);
    }

    // ==================== Sheet 1/2: 明細 ====================

    private void writeDetailSheet(Workbook wb, ExcelStyleHelper styles, String sheetName,
                                  int year,
                                  Map<Integer, List<CompanyMonthData>> data,
                                  Map<Integer, CompanyMonthData> subtotals,
                                  List<String> companyList,
                                  boolean isCumulative, int maxPeriod) {
        Sheet sheet = wb.createSheet(sheetName);
        int rowIdx = 0;

        // Row 0: 險種代號
        Row row0 = sheet.createRow(rowIdx++);
        createCell(row0, 2, "險種代號", styles.getHeaderStyle());
        List<String> codes = CategoryMapping.ALL_INSURANCE_CODES;
        for (int i = 0; i < codes.size(); i++) {
            createCell(row0, 3 + i, codes.get(i), styles.getHeaderStyle());
        }

        // Row 1: 表頭
        Row row1 = sheet.createRow(rowIdx++);
        createCell(row1, 0, "代號", styles.getHeaderStyle());
        createCell(row1, 1, "月份", styles.getHeaderStyle());
        createCell(row1, 2, "公司別/險種", styles.getHeaderStyle());
        for (int i = 0; i < codes.size(); i++) {
            String shortName = CategoryMapping.CODE_TO_SHORT_NAME.getOrDefault(codes.get(i), codes.get(i));
            createCell(row1, 3 + i, shortName, styles.getSubHeaderStyle());
        }
        createCell(row1, 3 + codes.size(), "合計", styles.getHeaderStyle());

        // 資料列
        long[] grandTotal = new long[codes.size() + 1]; // 總計

        for (int period = 1; period <= maxPeriod; period++) {
            List<CompanyMonthData> periodData = data.get(period);
            if (periodData == null) continue;

            String periodLabel = isCumulative ? CumulativeCalculator.getPeriodLabel(period) : String.valueOf(period);

            // 各公司
            for (CompanyMonthData cd : periodData) {
                Row row = sheet.createRow(rowIdx++);
                createCell(row, 0, cd.getCompanyCode(), styles.getCompanyStyle());
                createCell(row, 1, periodLabel, styles.getCompanyStyle());
                createCell(row, 2, cd.getCompanyName(), styles.getCompanyStyle());
                for (int i = 0; i < codes.size(); i++) {
                    createCell(row, 3 + i, cd.getPremium(codes.get(i)), styles.getNumberStyle());
                }
                createCell(row, 3 + codes.size(), cd.getTotal(), styles.getNumberStyle());
            }

            // 小計列
            CompanyMonthData subtotal = subtotals.get(period);
            if (subtotal != null) {
                Row stRow = sheet.createRow(rowIdx++);
                createCell(stRow, 0, "", styles.getSubtotalStyle());
                createCell(stRow, 1, periodLabel, styles.getSubtotalStyle());
                createCell(stRow, 2, "小計", styles.getSubtotalStyle());
                for (int i = 0; i < codes.size(); i++) {
                    long val = subtotal.getPremium(codes.get(i));
                    createCell(stRow, 3 + i, val, styles.getSubtotalNumberStyle());
                    grandTotal[i] += val;
                }
                long stTotal = subtotal.getTotal();
                createCell(stRow, 3 + codes.size(), stTotal, styles.getSubtotalNumberStyle());
                grandTotal[codes.size()] += stTotal;
            }
        }

        // 總計列 (僅單月表，累計表不需要)
        if (!isCumulative) {
            Row totalRow = sheet.createRow(rowIdx);
            createCell(totalRow, 0, "", styles.getSubtotalStyle());
            createCell(totalRow, 1, "總計", styles.getSubtotalStyle());
            createCell(totalRow, 2, "", styles.getSubtotalStyle());
            for (int i = 0; i < codes.size(); i++) {
                createCell(totalRow, 3 + i, grandTotal[i], styles.getSubtotalNumberStyle());
            }
            createCell(totalRow, 3 + codes.size(), grandTotal[codes.size()], styles.getSubtotalNumberStyle());
        }

        // 自動欄寬
        for (int c = 0; c <= 3 + codes.size(); c++) {
            sheet.setColumnWidth(c, c <= 2 ? 4000 : 3500);
        }
    }

    // ==================== Sheet 3/4: 總表 ====================

    private void writeSummarySheet(Workbook wb, ExcelStyleHelper styles, String sheetName,
                                   int year, int latestMonth,
                                   Map<Integer, List<CompanyMonthData>> data,
                                   Map<Integer, CompanyMonthData> subtotals,
                                   Map<Integer, CompanyMonthData> priorSubtotals,
                                   List<String> companyList,
                                   boolean isCumulative, int maxPeriod) {
        Sheet sheet = wb.createSheet(sheetName);
        List<SubCategory> subCats = categoryCalculator.getOutputSubCategories();
        int priorYear = year - 1;
        int rowIdx = 0;

        // Row 0: 標題
        String titleText = isCumulative
                ? String.format("%d年度 1-%d 月份各保險公司簽單保費統計累計總表", year, latestMonth)
                : String.format("%d年度 %d 月份各保險公司簽單保費統計總表", year, latestMonth);
        Row titleRow = sheet.createRow(rowIdx++);
        createCell(titleRow, 0, titleText, styles.getTitleStyle());
        int lastDataCol = subCats.size() + 3 + 2; // D~subCats + T(合計) + U(去年) + V(成長率)
        mergeRegion(sheet, 0, 0, 0, lastDataCol);

        // Row 1: 單位
        Row unitRow = sheet.createRow(rowIdx++);
        createCell(unitRow, lastDataCol, "單位:新台幣元", styles.getSubHeaderStyle());

        // Row 2-4: 表頭 (簡化版)
        Row headerRow1 = sheet.createRow(rowIdx++);
        Row headerRow2 = sheet.createRow(rowIdx++);
        Row headerRow3 = sheet.createRow(rowIdx++);

        createCell(headerRow2, 0, "代號", styles.getHeaderStyle());
        createCell(headerRow2, 1, "月份", styles.getHeaderStyle());
        createCell(headerRow2, 2, "公司別/險種", styles.getHeaderStyle());
        mergeRegion(sheet, 2, 4, 0, 0);
        mergeRegion(sheet, 2, 4, 1, 1);
        mergeRegion(sheet, 2, 4, 2, 2);

        // 子分類表頭
        int col = 3;
        for (SubCategory sub : subCats) {
            createCell(headerRow2, col, sub.getName(), styles.getHeaderStyle());
            mergeRegion(sheet, 2, 4, col, col);
            col++;
        }

        // T: 今年合計, U: 去年合計, V: 成長率
        int tCol = col;
        int uCol = col + 1;
        int vCol = col + 2;

        createCell(headerRow1, tCol, "比較", styles.getHeaderStyle());
        mergeRegion(sheet, 2, 2, tCol, vCol);
        createCell(headerRow2, tCol, year + "年度", styles.getHeaderStyle());
        createCell(headerRow2, uCol, priorYear + "年度", styles.getHeaderStyle());
        createCell(headerRow2, vCol, "成長率", styles.getHeaderStyle());
        createCell(headerRow3, tCol, "合計", styles.getSubHeaderStyle());
        createCell(headerRow3, uCol, "合計", styles.getSubHeaderStyle());
        createCell(headerRow3, vCol, "%", styles.getSubHeaderStyle());

        // 資料列
        long[] grandTotal = new long[subCats.size() + 3];

        for (int period = 1; period <= maxPeriod; period++) {
            List<CompanyMonthData> periodData = data.get(period);
            if (periodData == null) continue;

            String periodLabel = isCumulative ? CumulativeCalculator.getPeriodLabel(period) : String.valueOf(period);

            for (CompanyMonthData cd : periodData) {
                Row row = sheet.createRow(rowIdx++);
                createCell(row, 0, cd.getCompanyCode(), styles.getCompanyStyle());
                createCell(row, 1, periodLabel, styles.getCompanyStyle());
                createCell(row, 2, cd.getCompanyName(), styles.getCompanyStyle());

                Map<String, Long> catAmounts = categoryCalculator.calculateSubCategories(cd);
                long currentTotal = categoryCalculator.calculateTotal(catAmounts);

                col = 3;
                for (SubCategory sub : subCats) {
                    createCell(row, col++, catAmounts.getOrDefault(sub.getName(), 0L), styles.getNumberStyle());
                }
                createCell(row, tCol, currentTotal, styles.getNumberStyle());
                // U 和 V 留空 (個別公司不比較去年)
                createCell(row, uCol, "", styles.getCompanyStyle());
                createCell(row, vCol, "", styles.getCompanyStyle());
            }

            // 小計列
            CompanyMonthData subtotal = subtotals.get(period);
            if (subtotal != null) {
                Row stRow = sheet.createRow(rowIdx++);
                createCell(stRow, 0, "", styles.getSubtotalStyle());
                createCell(stRow, 1, periodLabel, styles.getSubtotalStyle());
                createCell(stRow, 2, "小計", styles.getSubtotalStyle());

                Map<String, Long> catAmounts = categoryCalculator.calculateSubCategories(subtotal);
                long currentTotal = categoryCalculator.calculateTotal(catAmounts);

                // 去年小計
                CompanyMonthData priorSub = priorSubtotals != null ? priorSubtotals.get(period) : null;
                Map<String, Long> priorCatAmounts = priorSub != null
                        ? categoryCalculator.calculateSubCategories(priorSub) : Collections.emptyMap();
                long priorTotal = priorSub != null ? categoryCalculator.calculateTotal(priorCatAmounts) : 0;

                col = 3;
                for (int i = 0; i < subCats.size(); i++) {
                    long val = catAmounts.getOrDefault(subCats.get(i).getName(), 0L);
                    createCell(stRow, col++, val, styles.getSubtotalNumberStyle());
                    grandTotal[i] += val;
                }

                createCell(stRow, tCol, currentTotal, styles.getSubtotalNumberStyle());
                createCell(stRow, uCol, priorTotal, styles.getSubtotalNumberStyle());
                grandTotal[subCats.size()] += currentTotal;
                grandTotal[subCats.size() + 1] += priorTotal;

                // 成長率
                com.insurance.report.model.GrowthRateResult gr =
                        com.insurance.report.util.GrowthRateUtil.calculate(
                                currentTotal, priorTotal,
                                sheetName + "/" + periodLabel + "/小計");
                createCell(stRow, vCol, gr.getRate(), styles.getPercentStyle());
            }
        }

        // 總計列 (僅單月表)
        if (!isCumulative) {
            Row totalRow = sheet.createRow(rowIdx);
            createCell(totalRow, 0, "", styles.getSubtotalStyle());
            createCell(totalRow, 1, "總計", styles.getSubtotalStyle());
            createCell(totalRow, 2, "", styles.getSubtotalStyle());
            col = 3;
            for (int i = 0; i < subCats.size(); i++) {
                createCell(totalRow, col++, grandTotal[i], styles.getSubtotalNumberStyle());
            }
            createCell(totalRow, tCol, grandTotal[subCats.size()], styles.getSubtotalNumberStyle());
            createCell(totalRow, uCol, grandTotal[subCats.size() + 1], styles.getSubtotalNumberStyle());

            com.insurance.report.model.GrowthRateResult totalGr =
                    com.insurance.report.util.GrowthRateUtil.calculate(
                            grandTotal[subCats.size()], grandTotal[subCats.size() + 1],
                            sheetName + "/總計");
            createCell(totalRow, vCol, totalGr.getRate(), styles.getPercentStyle());
        }

        // 欄寬
        for (int c = 0; c <= lastDataCol; c++) {
            sheet.setColumnWidth(c, c <= 2 ? 4000 : 3800);
        }
    }

    // ==================== Sheet 5: 歸屬 ====================

    private void writeGuishuSheet(Workbook wb, ExcelStyleHelper styles) {
        Sheet sheet = wb.createSheet("歸屬");
        int rowIdx = 0;

        // 表頭
        Row header = sheet.createRow(rowIdx++);
        createCell(header, 0, "歸屬", styles.getHeaderStyle());
        createCell(header, 1, "險種代號", styles.getHeaderStyle());
        createCell(header, 2, "險種名稱", styles.getHeaderStyle());

        for (String majorCat : CategoryMapping.MAJOR_CATEGORIES) {
            List<String> codesInCategory = new ArrayList<>();
            for (SubCategory sub : CategoryMapping.SUB_CATEGORIES) {
                if (sub.getMajorCategory().equals(majorCat)) {
                    codesInCategory.addAll(sub.getInsuranceCodes());
                }
            }
            if (majorCat.equals("國外分進")) {
                codesInCategory.addAll(CategoryMapping.OVERSEAS_REINSURANCE.getInsuranceCodes());
            }

            int startRow = rowIdx;
            for (String code : codesInCategory) {
                Row row = sheet.createRow(rowIdx);
                createCell(row, 0, majorCat, styles.getCompanyStyle());
                createCell(row, 1, code, styles.getCompanyStyle());
                createCell(row, 2, CategoryMapping.CODE_TO_SHORT_NAME.getOrDefault(code, ""), styles.getCompanyStyle());
                rowIdx++;
            }
            if (codesInCategory.size() > 1) {
                mergeRegion(sheet, startRow, rowIdx - 1, 0, 0);
            }
        }

        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 3000);
        sheet.setColumnWidth(2, 6000);
    }

    // ==================== Sheet 6: 排名 ====================

    private void writeRankingSheet(Workbook wb, ExcelStyleHelper styles,
                                   int year, int latestMonth,
                                   List<RankingCalculator.RankEntry> ranking) {
        Sheet sheet = wb.createSheet("排名");
        int rowIdx = 0;

        // 標題
        Row titleRow = sheet.createRow(rowIdx++);
        String title = String.format("%d年度 1-%d月累計保費排名", year, latestMonth);
        createCell(titleRow, 0, title, styles.getTitleStyle());
        mergeRegion(sheet, 0, 0, 0, 2);

        // 表頭
        Row header = sheet.createRow(rowIdx++);
        createCell(header, 0, "排名", styles.getHeaderStyle());
        createCell(header, 1, "公司名稱", styles.getHeaderStyle());
        createCell(header, 2, "保費合計", styles.getHeaderStyle());

        // 資料
        for (RankingCalculator.RankEntry entry : ranking) {
            Row row = sheet.createRow(rowIdx++);
            createCell(row, 0, (long) entry.getRank(), styles.getCompanyStyle());
            createCell(row, 1, entry.getCompany().getName(), styles.getCompanyStyle());
            createCell(row, 2, entry.getTotalPremium(), styles.getNumberStyle());
        }

        sheet.setColumnWidth(0, 2000);
        sheet.setColumnWidth(1, 5000);
        sheet.setColumnWidth(2, 5000);
    }

    // --- 工具方法 ---

    private List<String> getCompanyList(Map<Integer, List<CompanyMonthData>> data) {
        Set<String> codes = new TreeSet<>();
        data.values().stream()
                .flatMap(Collection::stream)
                .forEach(d -> codes.add(d.getCompanyCode()));
        return new ArrayList<>(codes);
    }
}
