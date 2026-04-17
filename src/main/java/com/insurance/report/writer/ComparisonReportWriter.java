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

    private final CategoryMapping categoryMapping;

    /** 比較表使用的子分類 (不含國外分進)，在首次 write 時初始化 */
    private List<SubCategory> comparisonCategories;
    private List<String> categoryNames;

    public ComparisonReportWriter(CategoryMapping categoryMapping) {
        this.categoryMapping = categoryMapping;
    }

    private void ensureInitialized() {
        if (comparisonCategories == null) {
            comparisonCategories = categoryMapping.getSubCategories();
            categoryNames = new ArrayList<>();
            for (SubCategory sub : comparisonCategories) {
                categoryNames.add(sub.getName());
            }
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
        ensureInitialized();
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
        createCell(titleRow, 0, title, styles.getCmpTitleStyle());
        mergeRegion(sheet, 0, 0, 0, 15);

        // Row 1: 空白
        sheet.createRow(rowIdx++);

        // Row 2: 月份 & 單位
        Row periodRow = sheet.createRow(rowIdx++);
        createCell(periodRow, 0, month + "月份", styles.getCmpPeriodStyle());
        createCell(periodRow, 16, "      單位:元", styles.getCmpSubHeaderStyle());

        // Row 3-5: 表頭 (3 層)
        rowIdx = writeComparisonHeaders(sheet, styles, rowIdx);

        // Row 6: 空白分隔
        rowIdx++;

        // === 區段一：單月比較 (Row 7-11) ===
        int priorRow = rowIdx;    // row 7 (0-based 6)
        rowIdx = writeValueRow(sheet, styles, rowIdx,
                String.format("%d/%d", priorYear, month), monthly.getPriorValues(), true);

        int currentRow = rowIdx;  // row 8 (0-based 7)
        rowIdx = writeValueRow(sheet, styles, rowIdx,
                String.format("%d/%d", year, month), monthly.getCurrentValues(), false);

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
        createCell(cumTitleRow, 0, String.format("1-%d月累計數", month), styles.getCmpPeriodStyle());

        // Row 15-17: 累計表頭
        rowIdx = writeComparisonHeaders(sheet, styles, rowIdx);

        // Row 17: 空白
        rowIdx++;

        // === 區段二：累計比較 (Row 18-22) ===
        int cumPriorRow = rowIdx;
        rowIdx = writeValueRow(sheet, styles, rowIdx,
                String.format("%d/1-%d", priorYear, month), cumulative.getPriorValues(), true);

        int cumCurrentRow = rowIdx;
        rowIdx = writeValueRow(sheet, styles, rowIdx,
                String.format("%d/1-%d", year, month), cumulative.getCurrentValues(), false);

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
        for (int c = 1; c <= categoryNames.size() + 1; c++) {
            sheet.setColumnWidth(c, 4000);
        }
    }

    private int writeComparisonHeaders(Sheet sheet, ExcelStyleHelper styles, int startRow) {
        Row h1 = sheet.createRow(startRow);
        Row h2 = sheet.createRow(startRow + 1);
        Row h3 = sheet.createRow(startRow + 2);

        // A: 年/月 險種 (span 3 rows)
        createCell(h1, 0, "年/月   險種", styles.getCmpHeaderStyle());
        mergeRegion(sheet, startRow, startRow + 2, 0, 0);

        // 動態計算各大類在子分類清單中的起止位置
        int totalSubCats = comparisonCategories.size();
        String currentMajor = null;
        int groupStartIdx = 0;
        List<int[]> majorRanges = new ArrayList<>();
        List<String> majorNames = new ArrayList<>();
        for (int i = 0; i < totalSubCats; i++) {
            String major = comparisonCategories.get(i).getMajorCategory();
            if (!major.equals(currentMajor)) {
                if (currentMajor != null) {
                    majorRanges.add(new int[]{groupStartIdx, i - 1});
                    majorNames.add(currentMajor);
                }
                currentMajor = major;
                groupStartIdx = i;
            }
        }
        if (currentMajor != null) {
            majorRanges.add(new int[]{groupStartIdx, totalSubCats - 1});
            majorNames.add(currentMajor);
        }

        int firstDataCol = 1; // 比較表從 B 欄開始

        for (int g = 0; g < majorRanges.size(); g++) {
            int[] range = majorRanges.get(g);
            String majorName = majorNames.get(g);
            int startCol = firstDataCol + range[0];
            int endCol = firstDataCol + range[1];
            int subCount = range[1] - range[0] + 1;

            if (subCount == 1) {
                // 單一子分類：合併 3 行
                createCell(h1, startCol, majorName, styles.getCmpHeaderStyle());
                mergeRegion(sheet, startRow, startRow + 2, startCol, startCol);
            } else {
                // 多子分類：h1 水平合併
                createCell(h1, startCol, majorName, styles.getCmpHeaderStyle());
                mergeRegion(sheet, startRow, startRow, startCol, endCol);

                // h2-h3: 子分類標題
                int i = range[0];
                while (i <= range[1]) {
                    SubCategory sub = comparisonCategories.get(i);
                    int col = firstDataCol + i;

                    if (sub.getHeaderGroup() != null) {
                        int hgStart = i;
                        String hg = sub.getHeaderGroup();
                        while (i <= range[1] && hg.equals(comparisonCategories.get(i).getHeaderGroup())) {
                            i++;
                        }
                        int hgEnd = i - 1;
                        int hgStartCol = firstDataCol + hgStart;
                        int hgEndCol = firstDataCol + hgEnd;

                        createCell(h2, hgStartCol, hg, styles.getCmpHeaderStyle());
                        if (hgEnd > hgStart) {
                            mergeRegion(sheet, startRow + 1, startRow + 1, hgStartCol, hgEndCol);
                        }
                        for (int j = hgStart; j <= hgEnd; j++) {
                            String label = comparisonCategories.get(j).getHeaderLabel();
                            if (label == null) label = comparisonCategories.get(j).getName();
                            createCell(h3, firstDataCol + j, label, styles.getCmpSubHeaderStyle());
                        }
                    } else if (sub.getShortHeader() != null && sub.getSubHeader() != null) {
                        createCell(h2, col, sub.getShortHeader(), styles.getCmpHeaderStyle());
                        createCell(h3, col, sub.getSubHeader(), styles.getCmpSubHeaderStyle());
                        i++;
                    } else {
                        createCell(h2, col, sub.getName(), styles.getCmpHeaderStyle());
                        mergeRegion(sheet, startRow + 1, startRow + 2, col, col);
                        i++;
                    }
                }
            }
        }

        // 合計欄
        int totalCol = firstDataCol + totalSubCats;
        createCell(h1, totalCol, "合計", styles.getCmpHeaderStyle());
        mergeRegion(sheet, startRow, startRow + 1, totalCol, totalCol);

        // 填充邊框
        styles.fillBorders(sheet, startRow, startRow + 2, 0, totalCol, styles.getCmpHeaderStyle());

        return startRow + 3;
    }

    /**
     * 寫入數值列 (去年/今年金額)：值直接寫入 B-P，合計 Q = SUM(B:P) 公式
     * @param isPrior true=去年(藍字), false=今年(黑字)
     */
    private int writeValueRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                              String label, Map<String, Long> values, boolean isPrior) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCmpLabelStyle());
        CellStyle numStyle = isPrior ? styles.getCmpBlueNumberStyle() : styles.getCmpNumberStyle();

        int col = 1;
        for (String catName : categoryNames) {
            createCell(row, col, values.getOrDefault(catName, 0L), numStyle);
            col++;
        }
        // Q 合計 = SUM(B:P)
        String sumFormula = String.format("SUM(%s:%s)",
                cellRef(1, rowIdx), cellRef(categoryNames.size(), rowIdx));
        createFormulaCell(row, col, sumFormula, numStyle);
        return rowIdx + 1;
    }

    /**
     * 寫入佔比列：公式 = 各 cell / $Q$currentRow，合計 Q = SUM(B:P)
     */
    private int writeProportionRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                                    String label, int currentValueRowIdx) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCmpRedLabelStyle());
        String totalRef = "$" + colLetter(categoryNames.size() + 1) + "$" + (currentValueRowIdx + 1);
        for (int col = 1; col <= categoryNames.size(); col++) {
            String formula = cellRef(col, currentValueRowIdx) + "/" + totalRef;
            createFormulaCell(row, col, formula, styles.getCmpPercentStyle());
        }
        // Q = SUM(B:P)
        String sumFormula = String.format("SUM(%s:%s)",
                cellRef(1, rowIdx), cellRef(categoryNames.size(), rowIdx));
        createFormulaCell(row, categoryNames.size() + 1, sumFormula, styles.getCmpPercentStyle());
        return rowIdx + 1;
    }

    /**
     * 寫入增減($)列：公式 = 今年 - 去年，合計 Q = Q今年 - Q去年
     */
    private int writeDifferenceRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                                    String label, int currentRowIdx, int priorRowIdx) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCmpLabelStyle());
        for (int col = 1; col <= categoryNames.size() + 1; col++) {
            String formula = cellRef(col, currentRowIdx) + "-" + cellRef(col, priorRowIdx);
            createFormulaCell(row, col, formula, styles.getCmpNumberStyle());
        }
        return rowIdx + 1;
    }

    /**
     * 寫入增減(%)列：公式 = IF(去年<0, NA(), 增減/ABS(去年))
     */
    private int writeGrowthRateRow(Sheet sheet, ExcelStyleHelper styles, int rowIdx,
                                    String label, int priorRowIdx, int diffRowIdx) {
        Row row = sheet.createRow(rowIdx);
        createCell(row, 0, label, styles.getCmpLabelStyle());
        for (int col = 1; col <= categoryNames.size() + 1; col++) {
            String priorRef = cellRef(col, priorRowIdx);
            String diffRef = cellRef(col, diffRowIdx);
            String formula = String.format("IF(%s<0,NA(),%s/ABS(%s))", priorRef, diffRef, priorRef);
            createFormulaCell(row, col, formula, styles.getCmpBoldPercentStyle());
        }
        return rowIdx + 1;
    }

    // ==================== Sheet 2: 增減原因 ====================

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
        createCell(titleRow, 0, title, styles.getCmpTitleStyle());
        mergeRegion(sheet, 0, 0, 0, 4);

        // Row 1: 空白
        rowIdx++;

        // Row 2: 表頭 (險種 A:B 合併 / 月份 C / 成長率 D / 增減原因 E)
        Row header = sheet.createRow(rowIdx++);
        createCell(header, 0, "險種", styles.getCmpHeaderStyle());
        mergeRegion(sheet, rowIdx - 1, rowIdx - 1, 0, 1);
        createCell(header, 2, "月份", styles.getCmpHeaderStyle());
        createCell(header, 3, "成長率", styles.getCmpHeaderStyle());
        createCell(header, 4, "增減原因", styles.getCmpHeaderStyle());

        int dataStartRow = rowIdx; // row 3 (0-based)

        // 動態建立增減原因分類結構
        List<String[]> reasonCategories = new ArrayList<>(); // [majorLabel, subLabel, colLetter]
        List<int[]> reasonMajorGroups = new ArrayList<>();   // [startIdx, endIdx exclusive]
        String curMajor = null;
        int majorStart = 0;

        for (int i = 0; i < comparisonCategories.size(); i++) {
            SubCategory sub = comparisonCategories.get(i);
            String major = sub.getMajorCategory();

            // 大類邊界
            if (!major.equals(curMajor)) {
                if (curMajor != null) {
                    reasonMajorGroups.add(new int[]{majorStart, i});
                }
                curMajor = major;
                majorStart = i;
            }

            // 標籤
            String majorLabel = "";
            if (i == majorStart || !major.equals(comparisonCategories.get(i - 1).getMajorCategory())) {
                majorLabel = major;
            }
            // 多子分類大類才顯示子類名
            long count = comparisonCategories.stream()
                    .filter(s -> s.getMajorCategory().equals(major)).count();
            String subLabel = count > 1 ? sub.getName() : "";
            String compColLetter = colLetter(i + 1);

            reasonCategories.add(new String[]{majorLabel, subLabel, compColLetter});
        }
        if (curMajor != null) {
            reasonMajorGroups.add(new int[]{majorStart, comparisonCategories.size()});
        }

        // 第一對使用文字值，後續用公式引用
        String cumPeriod = String.format("1-%d月", month);
        String monPeriod = String.format("%d月", month);
        int firstCumRow = dataStartRow;     // 第一個累計列的 0-based 行號
        int firstMonRow = dataStartRow + 1; // 第一個單月列的 0-based 行號

        // 比較增減率 sheet 中的成長率列 (1-based, 供 cross-sheet formula)
        int cumGrowthExcelRow = cumulativeGrowthRateRow + 1;
        int monGrowthExcelRow = monthlyGrowthRateRow + 1;

        boolean isFirst = true;
        for (int catIdx = 0; catIdx < reasonCategories.size(); catIdx++) {
            String[] cat = reasonCategories.get(catIdx);
            String majorGroup = cat[0];
            String subGroup = cat[1];
            String compColLetter = colLetter(catIdx + 1); // B, C, D, ..., P

            // 累計列
            Row cumRow = sheet.createRow(rowIdx++);
            createCell(cumRow, 0, majorGroup, styles.getCmpLabelStyle());
            createCell(cumRow, 1, subGroup, styles.getCmpLabelStyle());
            if (isFirst) {
                createCell(cumRow, 2, cumPeriod, styles.getCmpLabelStyle());
            } else {
                createFormulaCell(cumRow, 2, cellRef(2, firstCumRow), styles.getCmpLabelStyle());
            }
            // D: 成長率 = 比較增減率!{col}{cumGrowthRow}
            String cumFormula = String.format("比較增減率!%s%d", compColLetter, cumGrowthExcelRow);
            createFormulaCell(cumRow, 3, cumFormula, styles.getCmpPercentStyle());
            createCell(cumRow, 4, "", styles.getCmpLabelStyle());

            // 單月列
            Row monRow = sheet.createRow(rowIdx++);
            createCell(monRow, 0, "", styles.getCmpLabelStyle());
            createCell(monRow, 1, "", styles.getCmpLabelStyle());
            if (isFirst) {
                createCell(monRow, 2, monPeriod, styles.getCmpLabelStyle());
            } else {
                createFormulaCell(monRow, 2, cellRef(2, firstMonRow), styles.getCmpLabelStyle());
            }
            // D: 成長率 = 比較增減率!{col}{monGrowthRow}
            String monFormula = String.format("比較增減率!%s%d", compColLetter, monGrowthExcelRow);
            createFormulaCell(monRow, 3, monFormula, styles.getCmpPercentStyle());
            createCell(monRow, 4, "", styles.getCmpLabelStyle());

            isFirst = false;
        }

        // 合併大類儲存格
        for (int[] group : reasonMajorGroups) {
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
