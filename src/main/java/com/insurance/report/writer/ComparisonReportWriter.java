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
        Row h3 = sheet.createRow(startRow + 2);

        // A: 年/月 險種 (span 3 rows)
        createCell(h1, 0, "年/月   險種", styles.getHeaderStyle());
        mergeRegion(sheet, startRow, startRow + 2, 0, 0);

        // B=火險(1), C=水險(2), D=航空險(3): 獨立跨 3 行
        createCell(h1, 1, "火險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow, startRow + 2, 1, 1);
        createCell(h1, 2, "水險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow, startRow + 2, 2, 2);
        createCell(h1, 3, "航空險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow, startRow + 2, 3, 3);

        // E-I=汽車險 (cols 4-8): 合併 row h1
        createCell(h1, 4, "汽車險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow, startRow, 4, 8);
        // E=車體損失保險(4), F=任意責任險(5): 跨 h2-h3
        createCell(h2, 4, "車體損失保險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow + 1, startRow + 2, 4, 4);
        createCell(h2, 5, "任意責任險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow + 1, startRow + 2, 5, 5);
        // G-I=強制責任險 (cols 6-8): 合併 row h2
        createCell(h2, 6, "強制責任險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow + 1, startRow + 1, 6, 8);
        // G=汽車, H=機車, I=電動二輪車: row h3
        createCell(h3, 6, "汽車", styles.getSubHeaderStyle());
        createCell(h3, 7, "機車", styles.getSubHeaderStyle());
        createCell(h3, 8, "電動二輪車", styles.getSubHeaderStyle());

        // J-M=意外險 (cols 9-12): 合併 row h1
        createCell(h1, 9, "意外險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow, startRow, 9, 12);
        // J=責任險(9), K=工程險(10), L=信用保證(11): 跨 h2-h3
        createCell(h2, 9, "責任險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow + 1, startRow + 2, 9, 9);
        createCell(h2, 10, "工程險", styles.getHeaderStyle());
        mergeRegion(sheet, startRow + 1, startRow + 2, 10, 10);
        createCell(h2, 11, "信用保證", styles.getHeaderStyle());
        mergeRegion(sheet, startRow + 1, startRow + 2, 11, 11);
        // M=其他財產(12) row h2 + 責任保險 row h3
        createCell(h2, 12, "其他財產", styles.getHeaderStyle());
        createCell(h3, 12, "責任保險", styles.getSubHeaderStyle());

        // N=傷害險(13), O=天災險(14), P=健康險(15): row h2 only
        createCell(h2, 13, "傷害險", styles.getHeaderStyle());
        createCell(h2, 14, "天災險", styles.getHeaderStyle());
        createCell(h2, 15, "健康險", styles.getHeaderStyle());

        // Q=合計(16): 跨 h1-h2
        createCell(h1, 16, "合計", styles.getHeaderStyle());
        mergeRegion(sheet, startRow, startRow + 1, 16, 16);

        // 填充邊框
        styles.fillBorders(sheet, startRow, startRow + 2, 0, 16);

        return startRow + 3;
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

    /**
     * 增減原因結構定義
     * [0]=大類名稱, [1]=子類顯示名, [2]=比較增減率欄位字母 (用於公式引用)
     */
    private static final String[][] REASON_CATEGORIES = {
            {"火險", "", "B"},
            {"水險", "", "C"},
            {"航空", "", "D"},
            {"汽車險", "車體損", "E"},
            {"", "任意車責", "F"},
            {"", "強制車", "G"},
            {"", "強制機", "H"},
            {"", "強制電動二輪", "I"},
            {"意外險", "責任險", "J"},
            {"", "工程險", "K"},
            {"", "信用保險", "L"},
            {"", "其他責任險", "M"},
            {"傷害險", "", "N"},
            {"天災險", "", "O"},
            {"健康險", "", "P"},
    };

    /** 增減原因的大類 A 欄合併區域 (起始分類索引, 結束分類索引 exclusive) */
    private static final int[][] REASON_MAJOR_GROUPS = {
            {0, 1},    // 火險
            {1, 2},    // 水險
            {2, 3},    // 航空
            {3, 8},    // 汽車險 (5 子類)
            {8, 12},   // 意外險 (4 子類)
            {12, 13},  // 傷害險
            {13, 14},  // 天災險
            {14, 15},  // 健康險
    };

    private void writeReasonSheet(Workbook wb, ExcelStyleHelper styles,
                                   int year, int priorYear, int month,
                                   ComparisonCalculator.ComparisonRow monthly,
                                   ComparisonCalculator.ComparisonRow cumulative) {
        Sheet sheet = wb.createSheet("增減原因");
        int rowIdx = 0;

        // Row 0: 標題 (跨 A:E)
        Row titleRow = sheet.createRow(rowIdx++);
        String title = String.format("%d年與%d年產物保險業保費同期比較統計表\n保費收入、成長率及其增減原因分析如后：",
                year, priorYear);
        createCell(titleRow, 0, title, styles.getTitleStyle());
        mergeRegion(sheet, 0, 0, 0, 4);

        // Row 1: 空白
        rowIdx++;

        // Row 2: 表頭 (險種 A:B 合併 / 月份 C / 成長率 D / 增減原因 E)
        Row header = sheet.createRow(rowIdx++);
        createCell(header, 0, "險種", styles.getHeaderStyle());
        mergeRegion(sheet, rowIdx - 1, rowIdx - 1, 0, 1);
        createCell(header, 2, "月份", styles.getHeaderStyle());
        createCell(header, 3, "成長率", styles.getHeaderStyle());
        createCell(header, 4, "增減原因", styles.getHeaderStyle());

        int dataStartRow = rowIdx; // row 3 (0-based)

        // 15 類 × 2 列 (累計 + 單月)
        String cumPeriod = String.format("1-%d月", month);
        String monPeriod = String.format("%d月", month);

        // 成長率公式引用行號: 比較增減率 sheet 中
        // 單月成長率在 row 11 (1-based), 累計成長率在 row 22 (1-based)
        // 但我們用計算值，因為行號可能不固定

        for (String[] cat : REASON_CATEGORIES) {
            String majorGroup = cat[0];
            String subGroup = cat[1];
            String catName = getCategoryNameForReason(cat);

            // 累計列
            GrowthRateResult cumGr = cumulative.getGrowthRates().get(catName);
            Row cumRow = sheet.createRow(rowIdx++);
            createCell(cumRow, 0, majorGroup, styles.getCompanyStyle());
            createCell(cumRow, 1, subGroup, styles.getCompanyStyle());
            createCell(cumRow, 2, cumPeriod, styles.getCompanyStyle());
            createCell(cumRow, 3, cumGr != null ? cumGr.getRate() : 0.0, styles.getPercentStyle());
            createCell(cumRow, 4, "", styles.getCompanyStyle());

            // 單月列
            GrowthRateResult monGr = monthly.getGrowthRates().get(catName);
            Row monRow = sheet.createRow(rowIdx++);
            createCell(monRow, 0, "", styles.getCompanyStyle());
            createCell(monRow, 1, "", styles.getCompanyStyle());
            createCell(monRow, 2, monPeriod, styles.getCompanyStyle());
            createCell(monRow, 3, monGr != null ? monGr.getRate() : 0.0, styles.getPercentStyle());
            createCell(monRow, 4, "", styles.getCompanyStyle());
        }

        // 合併大類儲存格
        for (int[] group : REASON_MAJOR_GROUPS) {
            int startIdx = group[0];
            int endIdx = group[1];
            int mergeStartRow = dataStartRow + startIdx * 2;
            int mergeEndRow = dataStartRow + endIdx * 2 - 1;

            if (endIdx - startIdx == 1) {
                // 獨立大類 (火/水/航/傷/天/健): A:B 合併
                mergeRegion(sheet, mergeStartRow, mergeEndRow, 0, 1);
            } else {
                // 有子類的大類 (汽車/意外): A 欄合併
                mergeRegion(sheet, mergeStartRow, mergeEndRow, 0, 0);
                // 各子類 B 欄合併 (每子類 2 行)
                for (int i = startIdx; i < endIdx; i++) {
                    int subStart = dataStartRow + i * 2;
                    mergeRegion(sheet, subStart, subStart + 1, 1, 1);
                }
            }
        }

        sheet.setColumnWidth(0, 3500);
        sheet.setColumnWidth(1, 3500);
        sheet.setColumnWidth(2, 3000);
        sheet.setColumnWidth(3, 3500);
        sheet.setColumnWidth(4, 12000);
    }

    /**
     * 從 REASON_CATEGORIES 定義中找出對應的 CATEGORY_NAMES 名稱
     */
    private String getCategoryNameForReason(String[] cat) {
        String majorGroup = cat[0];
        String subGroup = cat[1];
        // 大類名稱對照
        if (subGroup.isEmpty()) {
            // 獨立大類: 直接用大類名 → 對應 CATEGORY_NAMES
            if (majorGroup.equals("航空")) return "航空險";
            return majorGroup;
        }
        // 子類名稱對照
        return switch (subGroup) {
            case "車體損" -> "車體損失險";
            case "任意車責" -> "任意責任險";
            case "強制車" -> "強制責任-汽車";
            case "強制機" -> "強制-機車";
            case "強制電動二輪" -> "強制-電動二輪";
            case "責任險" -> "責任險";
            case "工程險" -> "工程險";
            case "信用保險" -> "信用保證";
            case "其他責任險" -> "其他財產責任保險";
            default -> subGroup;
        };
    }
}
