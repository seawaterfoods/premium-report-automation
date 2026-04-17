package com.insurance.report.writer;

import com.insurance.report.model.CategoryMapping;
import com.insurance.report.model.CategoryMapping.SubCategory;
import com.insurance.report.calculator.ComparisonCalculator;
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

    // 寫入期間追蹤：比較增減率 sheet 中成長率列的 0-based 行號 (供增減原因公式引用)
    private int monthlyGrowthRateRow;
    private int cumulativeGrowthRateRow;

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
        int priorRow = rowIdx;    // row 7 (0-based 6)
        rowIdx = writeValueRow(sheet, styles, rowIdx,
                String.format("%d/%d", priorYear, month), monthly.getPriorValues());

        int currentRow = rowIdx;  // row 8 (0-based 7)
        rowIdx = writeValueRow(sheet, styles, rowIdx,
                String.format("%d/%d", year, month), monthly.getCurrentValues());

        // Row 9: 佔比 (引用今年列)
        rowIdx = writeProportionRow(sheet, styles, rowIdx,
                String.format("佔%d月比重", month), currentRow);

        // Row 10: 增減($) = 今年 - 去年
        int diffRow = rowIdx;
        rowIdx = writeDifferenceRow(sheet, styles, rowIdx,
                "同期增減($)", currentRow, priorRow);

        // Row 11: 增減(%) = IF(去年<0, NA(), 增減/ABS(去年))
        monthlyGrowthRateRow = rowIdx;
        rowIdx = writeGrowthRateRow(sheet, styles, rowIdx,
                "同期增減(%)", priorRow, diffRow);

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
        int cumPriorRow = rowIdx;
        rowIdx = writeValueRow(sheet, styles, rowIdx,
                String.format("%d/1-%d", priorYear, month), cumulative.getPriorValues());

        int cumCurrentRow = rowIdx;
        rowIdx = writeValueRow(sheet, styles, rowIdx,
                String.format("%d/1-%d", year, month), cumulative.getCurrentValues());

        // 佔累計比重
        rowIdx = writeProportionRow(sheet, styles, rowIdx,
                String.format("佔1-%d月累計比重", month), cumCurrentRow);

        // 累計增減($)
        int cumDiffRow = rowIdx;
        rowIdx = writeDifferenceRow(sheet, styles, rowIdx,
                "同期增減($)", cumCurrentRow, cumPriorRow);

        // 累計增減(%)
        cumulativeGrowthRateRow = rowIdx;
        writeGrowthRateRow(sheet, styles, rowIdx,
                "同期增減(%)", cumPriorRow, cumDiffRow);

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

    /**
     * 寫入數值列 (去年/今年金額)：值直接寫入 B-P，合計 Q = SUM(B:P) 公式
     */
    private int writeValueRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                              String label, Map<String, Long> values) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCompanyStyle());

        int col = 1;
        for (String catName : CATEGORY_NAMES) {
            createCell(row, col, values.getOrDefault(catName, 0L), styles.getNumberStyle());
            col++;
        }
        // Q 合計 = SUM(B:P)
        String sumFormula = String.format("SUM(%s:%s)",
                cellRef(1, rowIdx), cellRef(CATEGORY_NAMES.size(), rowIdx));
        createFormulaCell(row, col, sumFormula, styles.getNumberStyle());
        return rowIdx + 1;
    }

    /**
     * 寫入佔比列：公式 = 各 cell / $Q$currentRow，合計 Q = SUM(B:P)
     */
    private int writeProportionRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                                    String label, int currentValueRowIdx) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCompanyStyle());
        String totalRef = "$" + colLetter(CATEGORY_NAMES.size() + 1) + "$" + (currentValueRowIdx + 1);
        for (int col = 1; col <= CATEGORY_NAMES.size(); col++) {
            String formula = cellRef(col, currentValueRowIdx) + "/" + totalRef;
            createFormulaCell(row, col, formula, styles.getPercentStyle());
        }
        // Q = SUM(B:P)
        String sumFormula = String.format("SUM(%s:%s)",
                cellRef(1, rowIdx), cellRef(CATEGORY_NAMES.size(), rowIdx));
        createFormulaCell(row, CATEGORY_NAMES.size() + 1, sumFormula, styles.getPercentStyle());
        return rowIdx + 1;
    }

    /**
     * 寫入增減($)列：公式 = 今年 - 去年，合計 Q = Q今年 - Q去年
     */
    private int writeDifferenceRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                                    String label, int currentRowIdx, int priorRowIdx) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCompanyStyle());
        for (int col = 1; col <= CATEGORY_NAMES.size() + 1; col++) {
            String formula = cellRef(col, currentRowIdx) + "-" + cellRef(col, priorRowIdx);
            createFormulaCell(row, col, formula, styles.getNumberStyle());
        }
        return rowIdx + 1;
    }

    /**
     * 寫入增減(%)列：公式 = IF(去年<0, NA(), 增減/ABS(去年))
     */
    private int writeGrowthRateRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                                    String label, int priorRowIdx, int diffRowIdx) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCompanyStyle());
        for (int col = 1; col <= CATEGORY_NAMES.size() + 1; col++) {
            String priorRef = cellRef(col, priorRowIdx);
            String diffRef = cellRef(col, diffRowIdx);
            String formula = String.format("IF(%s<0,NA(),%s/ABS(%s))", priorRef, diffRef, priorRef);
            createFormulaCell(row, col, formula, styles.getPercentStyle());
        }
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

        // 第一對使用文字值，後續用公式引用
        String cumPeriod = String.format("1-%d月", month);
        String monPeriod = String.format("%d月", month);
        int firstCumRow = dataStartRow;     // 第一個累計列的 0-based 行號
        int firstMonRow = dataStartRow + 1; // 第一個單月列的 0-based 行號

        // 比較增減率 sheet 中的成長率列 (1-based, 供 cross-sheet formula)
        int cumGrowthExcelRow = cumulativeGrowthRateRow + 1;
        int monGrowthExcelRow = monthlyGrowthRateRow + 1;

        boolean isFirst = true;
        for (int catIdx = 0; catIdx < REASON_CATEGORIES.length; catIdx++) {
            String[] cat = REASON_CATEGORIES[catIdx];
            String majorGroup = cat[0];
            String subGroup = cat[1];
            String compColLetter = colLetter(catIdx + 1); // B, C, D, ..., P

            // 累計列
            Row cumRow = sheet.createRow(rowIdx++);
            createCell(cumRow, 0, majorGroup, styles.getCompanyStyle());
            createCell(cumRow, 1, subGroup, styles.getCompanyStyle());
            if (isFirst) {
                createCell(cumRow, 2, cumPeriod, styles.getCompanyStyle());
            } else {
                createFormulaCell(cumRow, 2, cellRef(2, firstCumRow), styles.getCompanyStyle());
            }
            // D: 成長率 = 比較增減率!{col}{cumGrowthRow}
            String cumFormula = String.format("比較增減率!%s%d", compColLetter, cumGrowthExcelRow);
            createFormulaCell(cumRow, 3, cumFormula, styles.getPercentStyle());
            createCell(cumRow, 4, "", styles.getCompanyStyle());

            // 單月列
            Row monRow = sheet.createRow(rowIdx++);
            createCell(monRow, 0, "", styles.getCompanyStyle());
            createCell(monRow, 1, "", styles.getCompanyStyle());
            if (isFirst) {
                createCell(monRow, 2, monPeriod, styles.getCompanyStyle());
            } else {
                createFormulaCell(monRow, 2, cellRef(2, firstMonRow), styles.getCompanyStyle());
            }
            // D: 成長率 = 比較增減率!{col}{monGrowthRow}
            String monFormula = String.format("比較增減率!%s%d", compColLetter, monGrowthExcelRow);
            createFormulaCell(monRow, 3, monFormula, styles.getPercentStyle());
            createCell(monRow, 4, "", styles.getCompanyStyle());

            isFirst = false;
        }

        // 合併大類儲存格
        for (int[] group : REASON_MAJOR_GROUPS) {
            int startIdx = group[0];
            int endIdx = group[1];
            int mergeStartRow = dataStartRow + startIdx * 2;
            int mergeEndRow = dataStartRow + endIdx * 2 - 1;

            if (endIdx - startIdx == 1) {
                mergeRegion(sheet, mergeStartRow, mergeEndRow, 0, 1);
            } else {
                mergeRegion(sheet, mergeStartRow, mergeEndRow, 0, 0);
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

}
