package com.insurance.report.writer;

import com.insurance.report.calculator.*;
import com.insurance.report.config.AppConfig;
import com.insurance.report.model.*;
import com.insurance.report.model.CategoryMapping.SubCategory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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
    private final CategoryMapping categoryMapping;

    public PremiumReportWriter(AppConfig config, CategoryCalculator categoryCalculator,
                               CategoryMapping categoryMapping) {
        this.config = config;
        this.categoryCalculator = categoryCalculator;
        this.categoryMapping = categoryMapping;
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
            Map<Integer, List<CompanyMonthData>> priorMonthlyData,
            Map<Integer, List<CompanyMonthData>> priorCumulativeData,
            Map<Integer, CompanyMonthData> priorMonthlySubtotals,
            Map<Integer, CompanyMonthData> priorCumulativeSubtotals,
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
                    priorMonthlyData, priorMonthlySubtotals,
                    companyList, false, 12);

            // Sheet 4: {YYY}總累
            writeSummarySheet(wb, styles, year + "總累", year, latestMonth,
                    cumulativeData, cumulativeSubtotals,
                    priorCumulativeData, priorCumulativeSubtotals,
                    companyList, true, latestMonth);

            // Sheet 5: 歸屬
            writeGuishuSheet(wb, styles);

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

        // Row 0: 險種代號 (排除 hidden-codes)
        Row row0 = sheet.createRow(rowIdx++);
        createCell(row0, 2, "險種代號", styles.getHeaderStyle());
        List<String> allCodes = categoryMapping.getAllInsuranceCodes();
        List<String> hiddenCodes = config.getColumns().getHiddenCodes();
        List<String> codes = new ArrayList<>();
        for (String code : allCodes) {
            if (!hiddenCodes.contains(code)) {
                codes.add(code);
            }
        }
        for (int i = 0; i < codes.size(); i++) {
            createCell(row0, 3 + i, codes.get(i), styles.getHeaderStyle());
        }

        // Row 1: 表頭
        Row row1 = sheet.createRow(rowIdx++);
        createCell(row1, 0, "代號", styles.getHeaderStyle());
        createCell(row1, 1, "月份", styles.getHeaderStyle());
        createCell(row1, 2, "公司別/險種", styles.getHeaderStyle());
        for (int i = 0; i < codes.size(); i++) {
            String shortName = categoryMapping.getCodeToShortName().getOrDefault(codes.get(i), codes.get(i));
            createCell(row1, 3 + i, shortName, styles.getSubHeaderStyle());
        }
        createCell(row1, 3 + codes.size(), "合計", styles.getHeaderStyle());

        int firstCodeCol = 3;
        int lastCodeCol = 3 + codes.size() - 1;
        int totalCol = 3 + codes.size();

        // 追蹤小計列行號 (供總計公式用)
        List<Integer> subtotalRows = new ArrayList<>();

        for (int period = 1; period <= maxPeriod; period++) {
            List<CompanyMonthData> periodData = data.get(period);
            if (periodData == null) continue;

            String periodLabel = isCumulative ? CumulativeCalculator.getPeriodLabel(period) : String.valueOf(period);
            int periodFirstRow = rowIdx;

            // 各公司
            for (CompanyMonthData cd : periodData) {
                Row row = sheet.createRow(rowIdx);
                createCell(row, 0, cd.getCompanyCode(), styles.getCompanyStyle());
                createCell(row, 1, periodLabel, styles.getMonthStyle());
                createCell(row, 2, cd.getCompanyName(), styles.getCompanyStyle());
                for (int i = 0; i < codes.size(); i++) {
                    createCell(row, 3 + i, cd.getPremium(codes.get(i)), styles.getNumberStyle());
                }
                // 合計 = SUM(D:AJ)
                String sumFormula = String.format("SUM(%s:%s)",
                        cellRef(firstCodeCol, rowIdx), cellRef(lastCodeCol, rowIdx));
                createFormulaCell(row, totalCol, sumFormula, styles.getNumberStyle());
                rowIdx++;
            }

            // 小計列：每欄 = SUM(該欄所有公司列)
            if (subtotals.get(period) != null && periodFirstRow < rowIdx) {
                Row stRow = sheet.createRow(rowIdx);
                createCell(stRow, 0, "", styles.getSubtotalStyle());
                createCell(stRow, 1, periodLabel, styles.getSubtotalStyle());
                createCell(stRow, 2, "小計", styles.getSubtotalStyle());
                for (int c = firstCodeCol; c <= totalCol; c++) {
                    String formula = String.format("SUM(%s:%s)",
                            cellRef(c, periodFirstRow), cellRef(c, rowIdx - 1));
                    createFormulaCell(stRow, c, formula, styles.getSubtotalNumberStyle());
                }
                subtotalRows.add(rowIdx);
                rowIdx++;
            }
        }

        // 總計列 (僅單月表)：每欄 = 所有小計列的 SUM
        if (!isCumulative && !subtotalRows.isEmpty()) {
            Row totalRow = sheet.createRow(rowIdx);
            createCell(totalRow, 0, "", styles.getSubtotalStyle());
            createCell(totalRow, 1, "總計", styles.getSubtotalStyle());
            createCell(totalRow, 2, "", styles.getSubtotalStyle());
            for (int c = firstCodeCol; c <= totalCol; c++) {
                StringBuilder formula = new StringBuilder();
                for (int i = 0; i < subtotalRows.size(); i++) {
                    if (i > 0) formula.append("+");
                    formula.append(cellRef(c, subtotalRows.get(i)));
                }
                createFormulaCell(totalRow, c, formula.toString(), styles.getSubtotalNumberStyle());
            }
        }

        // 自動欄寬
        for (int c = 0; c <= totalCol; c++) {
            sheet.setColumnWidth(c, c <= 2 ? 4000 : 3500);
        }

        // 凍結窗格: 固定前 2 列 (表頭) + 前 3 欄 (代號/月份/公司)
        sheet.createFreezePane(3, 2);

        // 自動篩選: 表頭列(row 1)至最後資料列
        int lastDataRow = sheet.getLastRowNum();
        sheet.setAutoFilter(new CellRangeAddress(1, lastDataRow, 0, totalCol));
    }

    private static final int FIRST_DATA_COL = 3;  // D 欄

    private void writeSummarySheet(Workbook wb, ExcelStyleHelper styles, String sheetName,
                                   int year, int latestMonth,
                                   Map<Integer, List<CompanyMonthData>> data,
                                   Map<Integer, CompanyMonthData> subtotals,
                                   Map<Integer, List<CompanyMonthData>> priorData,
                                   Map<Integer, CompanyMonthData> priorSubtotals,
                                   List<String> companyList,
                                   boolean isCumulative, int maxPeriod) {
        Sheet sheet = wb.createSheet(sheetName);
        int priorYear = year - 1;

        // 從設定檔動態計算子分類清單及欄位數
        List<SubCategory> allSummaryCats = categoryMapping.getAllSubCategoriesWithOverseas();
        int totalCats = allSummaryCats.size();

        // 固定欄位位置
        int tCol = FIRST_DATA_COL + totalCats;       // 今年合計
        int uCol = tCol + 1;                          // 去年合計
        int vCol = tCol + 2;                          // 成長率
        int lastCol = vCol;

        // === Row 0: 標題 ===
        String titleText = isCumulative
                ? String.format("%d年度 1-%d 月份各保險公司簽單保費統計累計總表", year, latestMonth)
                : String.format("%d年度 %d 月份各保險公司簽單保費統計總表", year, latestMonth);
        Row titleRow = sheet.createRow(0);
        createCell(titleRow, 0, titleText, styles.getTitleStyle());
        mergeRegion(sheet, 0, 0, 0, lastCol);

        // === Row 1: 單位 ===
        Row unitRow = sheet.createRow(1);
        createCell(unitRow, lastCol, "單位:新台幣元", styles.getSubHeaderStyle());

        // === Row 2-4: 動態表頭 ===
        Row groupRow = sheet.createRow(2);
        Row subRow = sheet.createRow(3);
        Row sub2Row = sheet.createRow(4);

        createCell(subRow, 0, "代號", styles.getHeaderStyle());
        createCell(subRow, 1, "月份", styles.getHeaderStyle());
        createCell(subRow, 2, "公司別/險種", styles.getHeaderStyle());

        // 計算各大類在子分類清單中的起止位置
        String currentMajor = null;
        int groupStartIdx = 0;
        List<int[]> majorRanges = new ArrayList<>();
        List<String> majorNames = new ArrayList<>();
        for (int i = 0; i < totalCats; i++) {
            String major = allSummaryCats.get(i).getMajorCategory();
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
            majorRanges.add(new int[]{groupStartIdx, totalCats - 1});
            majorNames.add(currentMajor);
        }

        for (int g = 0; g < majorRanges.size(); g++) {
            int[] range = majorRanges.get(g);
            String majorName = majorNames.get(g);
            int startCol = FIRST_DATA_COL + range[0];
            int endCol = FIRST_DATA_COL + range[1];
            int subCount = range[1] - range[0] + 1;

            if (subCount == 1) {
                // 單一子分類：合併 row 2-4
                createCell(groupRow, startCol, majorName, styles.getHeaderStyle());
                mergeRegion(sheet, 2, 4, startCol, startCol);
            } else {
                // 多子分類：row 2 水平合併
                createCell(groupRow, startCol, majorName, styles.getHeaderStyle());
                mergeRegion(sheet, 2, 2, startCol, endCol);

                // Row 3-4: 子分類標題
                int i = range[0];
                while (i <= range[1]) {
                    SubCategory sub = allSummaryCats.get(i);
                    int col = FIRST_DATA_COL + i;

                    if (sub.getHeaderGroup() != null) {
                        // 找出同 headerGroup 的連續子分類
                        int hgStart = i;
                        String hg = sub.getHeaderGroup();
                        while (i <= range[1] && hg.equals(allSummaryCats.get(i).getHeaderGroup())) {
                            i++;
                        }
                        int hgEnd = i - 1;
                        int hgStartCol = FIRST_DATA_COL + hgStart;
                        int hgEndCol = FIRST_DATA_COL + hgEnd;

                        createCell(subRow, hgStartCol, hg, styles.getHeaderStyle());
                        if (hgEnd > hgStart) {
                            mergeRegion(sheet, 3, 3, hgStartCol, hgEndCol);
                        }
                        for (int j = hgStart; j <= hgEnd; j++) {
                            String label = allSummaryCats.get(j).getHeaderLabel();
                            if (label == null) label = allSummaryCats.get(j).getName();
                            createCell(sub2Row, FIRST_DATA_COL + j, label, styles.getSubHeaderStyle());
                        }
                    } else if (sub.getShortHeader() != null && sub.getSubHeader() != null) {
                        // 兩行顯示
                        createCell(subRow, col, sub.getShortHeader(), styles.getHeaderStyle());
                        createCell(sub2Row, col, sub.getSubHeader(), styles.getSubHeaderStyle());
                        i++;
                    } else {
                        // 一般：合併 row 3-4
                        createCell(subRow, col, sub.getName(), styles.getHeaderStyle());
                        mergeRegion(sheet, 3, 4, col, col);
                        i++;
                    }
                }
            }
        }

        // 比較欄
        createCell(groupRow, lastCol, "比較", styles.getHeaderStyle());
        createCell(subRow, tCol, year + "年度", styles.getHeaderStyle());
        createCell(subRow, uCol, priorYear + "年度", styles.getHeaderStyle());
        createCell(subRow, vCol, "成長率", styles.getHeaderStyle());
        createCell(sub2Row, tCol, "合計", styles.getSubHeaderStyle());
        createCell(sub2Row, uCol, "合計", styles.getSubHeaderStyle());
        createCell(sub2Row, vCol, "%", styles.getSubHeaderStyle());

        // 填充表頭邊框
        styles.fillBorders(sheet, 2, 4, 0, lastCol);

        // === 資料列 (從 row 5 開始) ===
        int rowIdx = 5;

        // T 欄 SUM 範圍取決於是否含國外分進
        int totalSumEndCol = config.getColumns().isIncludeOverseasReinsurance()
                ? FIRST_DATA_COL + totalCats - 1   // 最後子分類欄
                : FIRST_DATA_COL + totalCats - 2;  // 倒數第二欄

        // 建立去年各期的公司代號 → CompanyMonthData 快速查詢表
        Map<Integer, Map<String, CompanyMonthData>> priorByPeriodAndCompany = new HashMap<>();
        if (priorData != null) {
            for (var entry : priorData.entrySet()) {
                Map<String, CompanyMonthData> byCompany = new HashMap<>();
                for (CompanyMonthData cd : entry.getValue()) {
                    byCompany.put(cd.getCompanyCode(), cd);
                }
                priorByPeriodAndCompany.put(entry.getKey(), byCompany);
            }
        }

        // 追蹤小計列行號 (供總計公式用)
        List<Integer> subtotalRows = new ArrayList<>();

        for (int period = 1; period <= maxPeriod; period++) {
            List<CompanyMonthData> periodData = data.get(period);
            if (periodData == null) continue;

            String periodLabel = isCumulative ? CumulativeCalculator.getPeriodLabel(period) : String.valueOf(period);
            int periodFirstRow = rowIdx;

            for (CompanyMonthData cd : periodData) {
                Row row = sheet.createRow(rowIdx);
                createCell(row, 0, cd.getCompanyCode(), styles.getCompanyStyle());
                createCell(row, 1, periodLabel, styles.getMonthStyle());
                createCell(row, 2, cd.getCompanyName(), styles.getCompanyStyle());

                Map<String, Long> catAmounts = categoryCalculator.calculateAllSubCategories(cd);

                int col = FIRST_DATA_COL;
                for (SubCategory sub : allSummaryCats) {
                    createCell(row, col++, catAmounts.getOrDefault(sub.getName(), 0L), styles.getNumberStyle());
                }
                // T 今年合計 = SUM(D:S 或 D:R)
                String tFormula = String.format("SUM(%s:%s)",
                        cellRef(FIRST_DATA_COL, rowIdx), cellRef(totalSumEndCol, rowIdx));
                createFormulaCell(row, tCol, tFormula, styles.getNumberStyle());

                // U 去年合計 (從去年同期同公司資料計算)
                Map<String, CompanyMonthData> priorPeriodMap = priorByPeriodAndCompany.get(period);
                CompanyMonthData priorCd = priorPeriodMap != null ? priorPeriodMap.get(cd.getCompanyCode()) : null;
                if (priorCd != null) {
                    Map<String, Long> priorCatAmounts = categoryCalculator.calculateAllSubCategories(priorCd);
                    long priorTotal = categoryCalculator.calculateOutputTotal(priorCatAmounts);
                    createCell(row, uCol, priorTotal, styles.getNumberStyle());
                    // V 成長率 = IF(U<=0, NA(), (T-U)/U)
                    String tRef = cellRef(tCol, rowIdx);
                    String uRef = cellRef(uCol, rowIdx);
                    String grFormula = String.format("IF(%s<=0,NA(),(%s-%s)/%s)", uRef, tRef, uRef, uRef);
                    createFormulaCell(row, vCol, grFormula, styles.getPercentStyle());
                } else {
                    createCell(row, uCol, "", styles.getCompanyStyle());
                    createCell(row, vCol, "", styles.getCompanyStyle());
                }
                rowIdx++;
            }

            // 小計列
            CompanyMonthData subtotal = subtotals.get(period);
            if (subtotal != null && periodFirstRow < rowIdx) {
                Row stRow = sheet.createRow(rowIdx);
                createCell(stRow, 0, "", styles.getSubtotalStyle());
                createCell(stRow, 1, periodLabel, styles.getSubtotalStyle());
                createCell(stRow, 2, "小計", styles.getSubtotalStyle());

                // 各子分類 = SUM(該欄所有公司列)
                for (int c = FIRST_DATA_COL; c < FIRST_DATA_COL + totalCats; c++) {
                    String formula = String.format("SUM(%s:%s)",
                            cellRef(c, periodFirstRow), cellRef(c, rowIdx - 1));
                    createFormulaCell(stRow, c, formula, styles.getSubtotalNumberStyle());
                }

                // T 今年合計 = SUM(D:S 或 D:R)
                String tFormula = String.format("SUM(%s:%s)",
                        cellRef(FIRST_DATA_COL, rowIdx), cellRef(totalSumEndCol, rowIdx));
                createFormulaCell(stRow, tCol, tFormula, styles.getSubtotalNumberStyle());

                // U 去年合計 = SUM(各公司 U 欄)
                String uSumFormula = String.format("SUM(%s:%s)",
                        cellRef(uCol, periodFirstRow), cellRef(uCol, rowIdx - 1));
                createFormulaCell(stRow, uCol, uSumFormula, styles.getSubtotalNumberStyle());

                // V 成長率 = IF(U<=0, NA(), (T-U)/U)
                String tRef = cellRef(tCol, rowIdx);
                String uRef = cellRef(uCol, rowIdx);
                String grFormula = String.format("IF(%s<=0,NA(),(%s-%s)/%s)", uRef, tRef, uRef, uRef);
                createFormulaCell(stRow, vCol, grFormula, styles.getPercentStyle());

                subtotalRows.add(rowIdx);
                rowIdx++;
            }
        }

        // 總計列 (僅單月表)
        if (!isCumulative && !subtotalRows.isEmpty()) {
            Row totalRow = sheet.createRow(rowIdx);
            createCell(totalRow, 0, "", styles.getSubtotalStyle());
            createCell(totalRow, 1, "總計", styles.getSubtotalStyle());
            createCell(totalRow, 2, "", styles.getSubtotalStyle());

            // 各子分類 + T + U = SUM of subtotal rows
            for (int c = FIRST_DATA_COL; c <= uCol; c++) {
                StringBuilder formula = new StringBuilder();
                for (int i = 0; i < subtotalRows.size(); i++) {
                    if (i > 0) formula.append("+");
                    formula.append(cellRef(c, subtotalRows.get(i)));
                }
                createFormulaCell(totalRow, c,
                        formula.toString(),
                        c <= FIRST_DATA_COL + totalCats - 1
                                ? styles.getSubtotalNumberStyle() : styles.getSubtotalNumberStyle());
            }

            // V 成長率 = IF(U<=0, NA(), (T-U)/U)
            String tRef = cellRef(tCol, rowIdx);
            String uRef = cellRef(uCol, rowIdx);
            String grFormula = String.format("IF(%s<=0,NA(),(%s-%s)/%s)", uRef, tRef, uRef, uRef);
            createFormulaCell(totalRow, vCol, grFormula, styles.getPercentStyle());
        }

        // 欄寬
        sheet.setColumnWidth(0, 2500);
        sheet.setColumnWidth(1, 2500);
        sheet.setColumnWidth(2, 5000);
        for (int c = FIRST_DATA_COL; c <= lastCol; c++) {
            sheet.setColumnWidth(c, 3800);
        }

        // 凍結窗格: 固定前 5 列 (標題+表頭) + 前 3 欄 (代號/月份/公司)
        sheet.createFreezePane(3, 5);
    }

    // ==================== Sheet 5: 歸屬 ====================

    private void writeGuishuSheet(Workbook wb, ExcelStyleHelper styles) {
        Sheet sheet = wb.createSheet("歸屬");

        // 收集各大類的險種代號列表
        List<String[]> categoryBlocks = new ArrayList<>();  // [majorCat, categoryNumber]
        List<List<String>> categoryCodeLists = new ArrayList<>();

        for (String majorCat : categoryMapping.getMajorCategories()) {
            List<String> codesInCategory = new ArrayList<>();
            for (SubCategory sub : categoryMapping.getSubCategories()) {
                if (sub.getMajorCategory().equals(majorCat)) {
                    codesInCategory.addAll(sub.getInsuranceCodes());
                }
            }
            if (majorCat.equals(categoryMapping.getOverseasReinsurance().getMajorCategory())) {
                codesInCategory.addAll(categoryMapping.getOverseasReinsurance().getInsuranceCodes());
            }
            String catNum = categoryMapping.getMajorCategoryNumber().getOrDefault(majorCat, "");
            categoryBlocks.add(new String[]{majorCat, catNum});
            categoryCodeLists.add(codesInCategory);
        }

        // 計算每個大類的起始/結束列 (0-based, 不含表頭)
        int totalDataRows = categoryCodeLists.stream().mapToInt(List::size).sum();
        int lastDataRow = totalDataRows; // 0-based (row 0=header, 1..N=data)

        // 計算每個大類的第一列索引 (1-based, 因為 row 0 是表頭)
        int[] categoryFirstRow = new int[categoryBlocks.size()];
        int[] categoryLastRow = new int[categoryBlocks.size()];
        int pos = 1;
        for (int i = 0; i < categoryBlocks.size(); i++) {
            categoryFirstRow[i] = pos;
            categoryLastRow[i] = pos + categoryCodeLists.get(i).size() - 1;
            pos += categoryCodeLists.get(i).size();
        }

        // 寫入表頭 (row 0)
        Row header = sheet.createRow(0);
        String[] headerTexts = {"類", "歸屬", "代號", "", "險種"};
        for (int c = 0; c < 5; c++) {
            BorderStyle left = (c == 0) ? BorderStyle.DOUBLE : BorderStyle.THIN;
            BorderStyle right = (c == 4) ? BorderStyle.DOUBLE : BorderStyle.THIN;
            CellStyle style = styles.getGuishuStyle(ExcelStyleHelper.GuishuFont.HEADER,
                    BorderStyle.DOUBLE, BorderStyle.NONE, left, right);
            createCell(header, c, headerTexts[c], style);
        }

        // 寫入資料列
        // 框線規則 (對照範例):
        //   每個大類邊界放一條 MEDIUM 線，放在上方 top 或下方 bottom (不重複)
        //   需要 bottom=MEDIUM: 單行大類 + 汽車(3) + 健康(7)
        //   需要 top=MEDIUM: 第一大類 + 前一大類沒有 bottom 的
        //   最後一列: bottom=DOUBLE
        Set<Integer> needsBottom = new HashSet<>();
        for (int i = 0; i < categoryBlocks.size(); i++) {
            if (categoryCodeLists.get(i).size() == 1 || i == 3 || i == 7) {
                needsBottom.add(i);
            }
        }
        // 前一大類是否有 bottom=MEDIUM (用來決定當前大類是否需要 top)
        boolean prevHadBottom = false;

        int rowIdx = 1;
        for (int catIdx = 0; catIdx < categoryBlocks.size(); catIdx++) {
            String majorCat = categoryBlocks.get(catIdx)[0];
            String catNum = categoryBlocks.get(catIdx)[1];
            List<String> codes = categoryCodeLists.get(catIdx);
            boolean catNeedsBottom = needsBottom.contains(catIdx);
            boolean firstInCategory = true;

            for (int codeIdx = 0; codeIdx < codes.size(); codeIdx++) {
                String code = codes.get(codeIdx);
                Row row = sheet.createRow(rowIdx);

                boolean isCatFirst = (codeIdx == 0);
                boolean isCatLast = (codeIdx == codes.size() - 1);
                boolean isTableLast = (rowIdx == lastDataRow);

                // 上框線
                BorderStyle topBorder = BorderStyle.NONE;
                if (isCatFirst && !prevHadBottom) {
                    topBorder = BorderStyle.MEDIUM;
                }

                // 下框線
                BorderStyle bottomBorder = BorderStyle.NONE;
                if (isTableLast) {
                    bottomBorder = BorderStyle.DOUBLE;
                } else if (isCatLast && catNeedsBottom) {
                    bottomBorder = BorderStyle.MEDIUM;
                }

                for (int c = 0; c < 5; c++) {
                    BorderStyle left = (c == 0) ? BorderStyle.DOUBLE : BorderStyle.THIN;
                    BorderStyle right = (c == 4) ? BorderStyle.DOUBLE : BorderStyle.THIN;

                    ExcelStyleHelper.GuishuFont fontType;
                    String cellValue;

                    switch (c) {
                        case 0: // 類序號
                            fontType = ExcelStyleHelper.GuishuFont.CATEGORY;
                            cellValue = firstInCategory ? catNum : "";
                            break;
                        case 1: // 歸屬
                            fontType = ExcelStyleHelper.GuishuFont.DATA;
                            cellValue = firstInCategory ? majorCat : "";
                            break;
                        case 2: // 代號
                            fontType = ExcelStyleHelper.GuishuFont.CODE;
                            cellValue = code;
                            break;
                        case 3: // 子分組
                            fontType = ExcelStyleHelper.GuishuFont.DATA;
                            cellValue = categoryMapping.getCodeToSubGroup().getOrDefault(code, "");
                            break;
                        default: // 險種全名
                            fontType = ExcelStyleHelper.GuishuFont.DATA;
                            cellValue = categoryMapping.getCodeToFullName().getOrDefault(code, "");
                            break;
                    }

                    CellStyle style = styles.getGuishuStyle(fontType, topBorder, bottomBorder, left, right);
                    createCell(row, c, cellValue, style);
                }

                firstInCategory = false;
                rowIdx++;
            }
            prevHadBottom = catNeedsBottom;
        }

        sheet.setColumnWidth(0, 2000);
        sheet.setColumnWidth(1, 3000);
        sheet.setColumnWidth(2, 2500);
        sheet.setColumnWidth(3, 3000);
        sheet.setColumnWidth(4, 8000);
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
