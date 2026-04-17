package com.insurance.report.writer;

import com.insurance.report.calculator.ComparisonCalculator;
import com.insurance.report.model.CategoryMapping;
import com.insurance.report.model.CategoryMapping.SubCategory;
import com.insurance.report.model.GrowthRateResult;
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
 * 同期比較分析表 Excel 輸出 (2 個頁簽)
 */
@Component
public class ComparisonReportWriter {

    private static final Logger log = LoggerFactory.getLogger(ComparisonReportWriter.class);

    /** 比較表使用的 15 類 (不含國外分進) */
    private static final List<SubCategory> COMPARISON_CATEGORIES;
    private static final List<String> CATEGORY_NAMES;

    static {
        COMPARISON_CATEGORIES = CategoryMapping.SUB_CATEGORIES;
        CATEGORY_NAMES = new ArrayList<>();
        for (SubCategory sub : COMPARISON_CATEGORIES) {
            CATEGORY_NAMES.add(sub.getName());
        }
    }

    /**
     * 產生同期比較分析表
     */
    public void write(
            int year, int month,
            ComparisonCalculator.ComparisonRow monthlyComparison,
            ComparisonCalculator.ComparisonRow cumulativeComparison,
            Path outputPath) throws IOException {

        log.info("產生同期比較分析表: {}", outputPath.getFileName());
        int priorYear = year - 1;

        try (Workbook wb = new XSSFWorkbook()) {
            ExcelStyleHelper styles = new ExcelStyleHelper(wb);

            writeComparisonSheet(wb, styles, year, priorYear, month,
                    monthlyComparison, cumulativeComparison);

            writeReasonSheet(wb, styles, year, priorYear, month,
                    monthlyComparison, cumulativeComparison);

            Files.createDirectories(outputPath.getParent());
            try (OutputStream os = Files.newOutputStream(outputPath)) {
                wb.write(os);
            }
        }

        log.info("同期比較分析表已產生: {}", outputPath);
    }

    // ==================== Sheet 1: 比較增減率 ====================

    private void writeComparisonSheet(Workbook wb, ExcelStyleHelper styles,
                                       int year, int priorYear, int month,
                                       ComparisonCalculator.ComparisonRow monthly,
                                       ComparisonCalculator.ComparisonRow cumulative) {
        Sheet sheet = wb.createSheet("比較增減率");
        int rowIdx = 0;

        // Row 0: 標題
        Row titleRow = sheet.createRow(rowIdx++);
        String title = String.format("中華民國%d年與%d年產物保險業保費同期比較統計表", year, priorYear);
        createCell(titleRow, 0, title, styles.getTitleStyle());
        mergeRegion(sheet, 0, 0, 0, 15);

        // Row 1: 空白
        sheet.createRow(rowIdx++);

        // Row 2: 月份 & 單位
        Row periodRow = sheet.createRow(rowIdx++);
        createCell(periodRow, 0, month + "月份", styles.getHeaderStyle());
        createCell(periodRow, 16, "      單位:元", styles.getSubHeaderStyle());

        // Row 3-5: 表頭 (3 層)
        rowIdx = writeComparisonHeaders(sheet, styles, rowIdx);

        // Row 6: 空白分隔
        rowIdx++;

        // === 區段一：單月比較 (Row 7-11) ===
        // Row 7: 去年同月
        rowIdx = writeComparisonDataRow(sheet, styles, rowIdx,
                String.format("%d/%d", priorYear, month),
                monthly.getPriorValues(), null, false);

        // Row 8: 今年同月
        rowIdx = writeComparisonDataRow(sheet, styles, rowIdx,
                String.format("%d/%d", year, month),
                monthly.getCurrentValues(), null, false);

        // Row 9: 佔月比重
        rowIdx = writeComparisonDataRow(sheet, styles, rowIdx,
                String.format("佔%d月比重", month),
                null, monthly.getProportions(), true);

        // Row 10: 同期增減($)
        rowIdx = writeComparisonDataRow(sheet, styles, rowIdx,
                "同期增減($)", monthly.getDifferences(), null, false);

        // Row 11: 同期增減(%)
        rowIdx = writeGrowthRateRow(sheet, styles, rowIdx,
                "同期增減(%)", monthly.getGrowthRates());

        // Row 12-13: 空白
        rowIdx += 2;

        // Row 14: 累計標題
        Row cumTitleRow = sheet.createRow(rowIdx++);
        createCell(cumTitleRow, 0, String.format("1-%d月累計數", month), styles.getHeaderStyle());

        // Row 15-17: 累計表頭
        rowIdx = writeComparisonHeaders(sheet, styles, rowIdx);

        // Row 17: 空白
        rowIdx++;

        // === 區段二：累計比較 (Row 18-22) ===
        // Row 18: 去年累計
        rowIdx = writeComparisonDataRow(sheet, styles, rowIdx,
                String.format("%d/1-%d", priorYear, month),
                cumulative.getPriorValues(), null, false);

        // Row 19: 今年累計
        rowIdx = writeComparisonDataRow(sheet, styles, rowIdx,
                String.format("%d/1-%d", year, month),
                cumulative.getCurrentValues(), null, false);

        // Row 20: 佔累計比重
        rowIdx = writeComparisonDataRow(sheet, styles, rowIdx,
                String.format("佔1-%d月累計比重", month),
                null, cumulative.getProportions(), true);

        // Row 21: 同期增減($)
        rowIdx = writeComparisonDataRow(sheet, styles, rowIdx,
                "同期增減($)", cumulative.getDifferences(), null, false);

        // Row 22: 同期增減(%)
        writeGrowthRateRow(sheet, styles, rowIdx,
                "同期增減(%)", cumulative.getGrowthRates());

        // 欄寬
        sheet.setColumnWidth(0, 5000);
        for (int c = 1; c <= 16; c++) {
            sheet.setColumnWidth(c, 4000);
        }
    }

    private int writeComparisonHeaders(Sheet sheet, ExcelStyleHelper styles, int startRow) {
        Row h1 = sheet.createRow(startRow);
        Row h2 = sheet.createRow(startRow + 1);

        createCell(h1, 0, "", styles.getHeaderStyle());

        // B~P: 15 類
        int col = 1;
        for (SubCategory sub : COMPARISON_CATEGORIES) {
            createCell(h1, col, sub.getName(), styles.getHeaderStyle());
            mergeRegion(sheet, startRow, startRow + 1, col, col);
            col++;
        }

        // Q: 合計
        createCell(h1, col, "合計", styles.getHeaderStyle());
        mergeRegion(sheet, startRow, startRow + 1, col, col);

        return startRow + 2;
    }

    private int writeComparisonDataRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                                        String label,
                                        Map<String, Long> longValues,
                                        Map<String, Double> doubleValues,
                                        boolean isPercent) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCompanyStyle());

        int col = 1;
        for (String catName : CATEGORY_NAMES) {
            if (isPercent && doubleValues != null) {
                createCell(row, col, doubleValues.getOrDefault(catName, 0.0), styles.getPercentStyle());
            } else if (longValues != null) {
                createCell(row, col, longValues.getOrDefault(catName, 0L), styles.getNumberStyle());
            }
            col++;
        }
        // 合計
        if (isPercent && doubleValues != null) {
            createCell(row, col, doubleValues.getOrDefault("合計", 1.0), styles.getPercentStyle());
        } else if (longValues != null) {
            createCell(row, col, longValues.getOrDefault("合計", 0L), styles.getNumberStyle());
        }

        return rowIdx + 1;
    }

    private int writeGrowthRateRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                                    String label,
                                    Map<String, GrowthRateResult> rates) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCompanyStyle());

        int col = 1;
        for (String catName : CATEGORY_NAMES) {
            GrowthRateResult gr = rates.get(catName);
            double rate = gr != null ? gr.getRate() : 0.0;
            createCell(row, col, rate, styles.getPercentStyle());
            col++;
        }
        // 合計
        GrowthRateResult totalGr = rates.get("合計");
        createCell(row, col, totalGr != null ? totalGr.getRate() : 0.0, styles.getPercentStyle());

        return rowIdx + 1;
    }

    // ==================== Sheet 2: 增減原因 ====================

    private void writeReasonSheet(Workbook wb, ExcelStyleHelper styles,
                                   int year, int priorYear, int month,
                                   ComparisonCalculator.ComparisonRow monthly,
                                   ComparisonCalculator.ComparisonRow cumulative) {
        Sheet sheet = wb.createSheet("增減原因");
        int rowIdx = 0;

        // 標題
        Row titleRow = sheet.createRow(rowIdx++);
        String title = String.format("中華民國%d年與%d年產物保險業保費同期比較增減原因", year, priorYear);
        createCell(titleRow, 0, title, styles.getTitleStyle());
        mergeRegion(sheet, 0, 0, 0, 4);

        // 空白
        rowIdx++;

        // 表頭
        Row header = sheet.createRow(rowIdx++);
        createCell(header, 0, "險種大類", styles.getHeaderStyle());
        createCell(header, 1, "險種子類", styles.getHeaderStyle());
        createCell(header, 2, "月份區間", styles.getHeaderStyle());
        createCell(header, 3, "成長率", styles.getHeaderStyle());
        createCell(header, 4, "增減原因", styles.getHeaderStyle());

        // 15 類 × 2 列 (累計 + 單月)
        // 大類分組: 火險, 水險, 航空, 汽車險(5子類), 意外險(4子類), 傷害險, 天災險, 健康險
        String[][] categories = {
                {"火險", ""},
                {"水險", ""},
                {"航空險", ""},
                {"車體損失險", "車體損"},
                {"任意責任險", "任意車責"},
                {"強制責任-汽車", "強制車"},
                {"強制-機車", "強制機"},
                {"強制-電動二輪", "強制電動二輪"},
                {"責任險", "責任險"},
                {"工程險", "工程險"},
                {"信用保證", "信用保險"},
                {"其他財產責任保險", "其他責任險"},
                {"傷害險", ""},
                {"天災險", ""},
                {"健康險", ""},
        };

        String[] majorGroups = {"火險", "水險", "航空險", "汽車險", "汽車險", "汽車險", "汽車險", "汽車險",
                "意外險", "意外險", "意外險", "意外險", "傷害險", "天災險", "健康險"};

        for (int i = 0; i < categories.length; i++) {
            String catName = categories[i][0];
            String displaySub = categories[i][1];
            String majorGroup = majorGroups[i];

            // 累計列
            GrowthRateResult cumGr = cumulative.getGrowthRates().get(catName);
            Row cumRow = sheet.createRow(rowIdx++);
            createCell(cumRow, 0, majorGroup, styles.getCompanyStyle());
            createCell(cumRow, 1, displaySub.isEmpty() ? "—" : displaySub, styles.getCompanyStyle());
            createCell(cumRow, 2, String.format("1-%d月", month), styles.getCompanyStyle());
            createCell(cumRow, 3, cumGr != null ? cumGr.getRate() : 0.0, styles.getPercentStyle());
            createCell(cumRow, 4, "", styles.getCompanyStyle()); // 增減原因留空

            // 單月列
            GrowthRateResult monGr = monthly.getGrowthRates().get(catName);
            Row monRow = sheet.createRow(rowIdx++);
            createCell(monRow, 0, "", styles.getCompanyStyle());
            createCell(monRow, 1, "", styles.getCompanyStyle());
            createCell(monRow, 2, String.format("%d月", month), styles.getCompanyStyle());
            createCell(monRow, 3, monGr != null ? monGr.getRate() : 0.0, styles.getPercentStyle());
            createCell(monRow, 4, "", styles.getCompanyStyle());
        }

        // 合併大類儲存格
        int dataStart = 3;
        int[] groupSizes = {2, 2, 2, 10, 8, 2, 2, 2}; // 火/水/航/汽車(5×2)/意外(4×2)/傷/天/健
        int pos = dataStart;
        for (int size : groupSizes) {
            if (size > 1) {
                mergeRegion(sheet, pos, pos + size - 1, 0, 0);
            }
            pos += size;
        }

        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 4000);
        sheet.setColumnWidth(2, 3500);
        sheet.setColumnWidth(3, 3500);
        sheet.setColumnWidth(4, 12000);
    }
}
