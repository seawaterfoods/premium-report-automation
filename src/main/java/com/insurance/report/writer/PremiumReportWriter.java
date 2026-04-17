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

    /** 總表固定使用全部 16 子分類 (含國外分進) */
    private static final List<SubCategory> ALL_SUMMARY_CATS = CategoryMapping.getAllSubCategoriesWithOverseas();
    private static final int FIRST_DATA_COL = 3;  // D 欄
    private static final int TOTAL_CATS = 16;      // 16 子分類

    // 子分類在 ALL_SUMMARY_CATS 中的索引範圍 (用於表頭分組)
    private static final int AUTO_GROUP_START = 3;   // 車體損失險
    private static final int AUTO_GROUP_END = 7;     // 電動二輪 (inclusive)
    private static final int ACCIDENT_GROUP_START = 8;  // 責任險
    private static final int ACCIDENT_GROUP_END = 11;   // 其他財產 (inclusive)
    private static final int MANDATORY_START = 5;    // 強制汽車
    private static final int MANDATORY_END = 7;      // 電動二輪 (inclusive)

    private void writeSummarySheet(Workbook wb, ExcelStyleHelper styles, String sheetName,
                                   int year, int latestMonth,
                                   Map<Integer, List<CompanyMonthData>> data,
                                   Map<Integer, CompanyMonthData> subtotals,
                                   Map<Integer, CompanyMonthData> priorSubtotals,
                                   List<String> companyList,
                                   boolean isCumulative, int maxPeriod) {
        Sheet sheet = wb.createSheet(sheetName);
        int priorYear = year - 1;

        // 固定欄位位置
        int tCol = FIRST_DATA_COL + TOTAL_CATS;     // T=19: 今年合計
        int uCol = tCol + 1;                          // U=20: 去年合計
        int vCol = tCol + 2;                          // V=21: 成長率
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

        // === Row 2: 分組標題 ===
        Row groupRow = sheet.createRow(2);
        // 火險/水險/航空 獨立跨 3 行 (row 2-4)
        createCell(groupRow, FIRST_DATA_COL + 0, "火險", styles.getHeaderStyle());
        mergeRegion(sheet, 2, 4, FIRST_DATA_COL + 0, FIRST_DATA_COL + 0);
        createCell(groupRow, FIRST_DATA_COL + 1, "水險", styles.getHeaderStyle());
        mergeRegion(sheet, 2, 4, FIRST_DATA_COL + 1, FIRST_DATA_COL + 1);
        createCell(groupRow, FIRST_DATA_COL + 2, "航空", styles.getHeaderStyle());
        mergeRegion(sheet, 2, 4, FIRST_DATA_COL + 2, FIRST_DATA_COL + 2);
        // 汽車險 合併 G:K (row 2)
        int autoStart = FIRST_DATA_COL + AUTO_GROUP_START;
        int autoEnd = FIRST_DATA_COL + AUTO_GROUP_END;
        createCell(groupRow, autoStart, "汽車險", styles.getHeaderStyle());
        mergeRegion(sheet, 2, 2, autoStart, autoEnd);
        // 意外險 合併 L:O (row 2)
        int accStart = FIRST_DATA_COL + ACCIDENT_GROUP_START;
        int accEnd = FIRST_DATA_COL + ACCIDENT_GROUP_END;
        createCell(groupRow, accStart, "意外險", styles.getHeaderStyle());
        mergeRegion(sheet, 2, 2, accStart, accEnd);
        // 比較 (row 2)
        createCell(groupRow, lastCol, "比較", styles.getHeaderStyle());

        // === Row 3: 子類標題 ===
        Row subRow = sheet.createRow(3);
        createCell(subRow, 0, "代號", styles.getHeaderStyle());
        createCell(subRow, 1, "月份", styles.getHeaderStyle());
        createCell(subRow, 2, "公司別/險種", styles.getHeaderStyle());
        // 汽車險子類
        createCell(subRow, FIRST_DATA_COL + 3, "車體損失險", styles.getHeaderStyle());
        mergeRegion(sheet, 3, 4, FIRST_DATA_COL + 3, FIRST_DATA_COL + 3);
        createCell(subRow, FIRST_DATA_COL + 4, "任意責任險", styles.getHeaderStyle());
        mergeRegion(sheet, 3, 4, FIRST_DATA_COL + 4, FIRST_DATA_COL + 4);
        int manStart = FIRST_DATA_COL + MANDATORY_START;
        int manEnd = FIRST_DATA_COL + MANDATORY_END;
        createCell(subRow, manStart, "強制責任險", styles.getHeaderStyle());
        mergeRegion(sheet, 3, 3, manStart, manEnd);
        // 意外險子類
        createCell(subRow, FIRST_DATA_COL + 8, "責任險", styles.getHeaderStyle());
        mergeRegion(sheet, 3, 4, FIRST_DATA_COL + 8, FIRST_DATA_COL + 8);
        createCell(subRow, FIRST_DATA_COL + 9, "工程險", styles.getHeaderStyle());
        mergeRegion(sheet, 3, 4, FIRST_DATA_COL + 9, FIRST_DATA_COL + 9);
        createCell(subRow, FIRST_DATA_COL + 10, "信用保證", styles.getHeaderStyle());
        mergeRegion(sheet, 3, 4, FIRST_DATA_COL + 10, FIRST_DATA_COL + 10);
        createCell(subRow, FIRST_DATA_COL + 11, "其他財產", styles.getHeaderStyle());
        // 獨立子類 (row 3 only, no vertical merge)
        createCell(subRow, FIRST_DATA_COL + 12, "傷害險", styles.getHeaderStyle());
        createCell(subRow, FIRST_DATA_COL + 13, "天災險", styles.getHeaderStyle());
        createCell(subRow, FIRST_DATA_COL + 14, "健康險", styles.getHeaderStyle());
        createCell(subRow, FIRST_DATA_COL + 15, "國外\n分進", styles.getHeaderStyle());
        mergeRegion(sheet, 3, 4, FIRST_DATA_COL + 15, FIRST_DATA_COL + 15);
        // 比較子類
        createCell(subRow, tCol, year + "年度", styles.getHeaderStyle());
        createCell(subRow, uCol, priorYear + "年度", styles.getHeaderStyle());
        createCell(subRow, vCol, "成長率", styles.getHeaderStyle());

        // === Row 4: 第三層標題 ===
        Row sub2Row = sheet.createRow(4);
        createCell(sub2Row, FIRST_DATA_COL + 5, "汽車", styles.getSubHeaderStyle());
        createCell(sub2Row, FIRST_DATA_COL + 6, "機車", styles.getSubHeaderStyle());
        createCell(sub2Row, FIRST_DATA_COL + 7, "電動二輪", styles.getSubHeaderStyle());
        createCell(sub2Row, FIRST_DATA_COL + 11, "責任保險", styles.getSubHeaderStyle());
        createCell(sub2Row, tCol, "合計", styles.getSubHeaderStyle());
        createCell(sub2Row, uCol, "合計", styles.getSubHeaderStyle());
        createCell(sub2Row, vCol, "%", styles.getSubHeaderStyle());

        // 填充表頭邊框
        styles.fillBorders(sheet, 2, 4, 0, lastCol);

        // === 資料列 (從 row 5 開始) ===
        int rowIdx = 5;
        long[] grandTotal = new long[TOTAL_CATS + 3];

        for (int period = 1; period <= maxPeriod; period++) {
            List<CompanyMonthData> periodData = data.get(period);
            if (periodData == null) continue;

            String periodLabel = isCumulative ? CumulativeCalculator.getPeriodLabel(period) : String.valueOf(period);

            for (CompanyMonthData cd : periodData) {
                Row row = sheet.createRow(rowIdx++);
                createCell(row, 0, cd.getCompanyCode(), styles.getCompanyStyle());
                createCell(row, 1, periodLabel, styles.getCompanyStyle());
                createCell(row, 2, cd.getCompanyName(), styles.getCompanyStyle());

                Map<String, Long> catAmounts = categoryCalculator.calculateAllSubCategories(cd);
                long currentTotal = categoryCalculator.calculateOutputTotal(catAmounts);

                int col = FIRST_DATA_COL;
                for (SubCategory sub : ALL_SUMMARY_CATS) {
                    createCell(row, col++, catAmounts.getOrDefault(sub.getName(), 0L), styles.getNumberStyle());
                }
                createCell(row, tCol, currentTotal, styles.getNumberStyle());
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

                Map<String, Long> catAmounts = categoryCalculator.calculateAllSubCategories(subtotal);
                long currentTotal = categoryCalculator.calculateOutputTotal(catAmounts);

                CompanyMonthData priorSub = priorSubtotals != null ? priorSubtotals.get(period) : null;
                Map<String, Long> priorCatAmounts = priorSub != null
                        ? categoryCalculator.calculateAllSubCategories(priorSub) : Collections.emptyMap();
                long priorTotal = priorSub != null ? categoryCalculator.calculateOutputTotal(priorCatAmounts) : 0;

                int col = FIRST_DATA_COL;
                for (int i = 0; i < ALL_SUMMARY_CATS.size(); i++) {
                    long val = catAmounts.getOrDefault(ALL_SUMMARY_CATS.get(i).getName(), 0L);
                    createCell(stRow, col++, val, styles.getSubtotalNumberStyle());
                    grandTotal[i] += val;
                }

                createCell(stRow, tCol, currentTotal, styles.getSubtotalNumberStyle());
                createCell(stRow, uCol, priorTotal, styles.getSubtotalNumberStyle());
                grandTotal[TOTAL_CATS] += currentTotal;
                grandTotal[TOTAL_CATS + 1] += priorTotal;

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
            int col = FIRST_DATA_COL;
            for (int i = 0; i < TOTAL_CATS; i++) {
                createCell(totalRow, col++, grandTotal[i], styles.getSubtotalNumberStyle());
            }
            createCell(totalRow, tCol, grandTotal[TOTAL_CATS], styles.getSubtotalNumberStyle());
            createCell(totalRow, uCol, grandTotal[TOTAL_CATS + 1], styles.getSubtotalNumberStyle());

            com.insurance.report.model.GrowthRateResult totalGr =
                    com.insurance.report.util.GrowthRateUtil.calculate(
                            grandTotal[TOTAL_CATS], grandTotal[TOTAL_CATS + 1],
                            sheetName + "/總計");
            createCell(totalRow, vCol, totalGr.getRate(), styles.getPercentStyle());
        }

        // 欄寬
        sheet.setColumnWidth(0, 2500);
        sheet.setColumnWidth(1, 2500);
        sheet.setColumnWidth(2, 5000);
        for (int c = FIRST_DATA_COL; c <= lastCol; c++) {
            sheet.setColumnWidth(c, 3800);
        }
    }

    // ==================== Sheet 5: 歸屬 ====================

    private void writeGuishuSheet(Workbook wb, ExcelStyleHelper styles) {
        Sheet sheet = wb.createSheet("歸屬");
        int rowIdx = 0;

        // 表頭: 類 / 歸屬 / 代號 / (子分組) / 險種
        Row header = sheet.createRow(rowIdx++);
        createCell(header, 0, "類", styles.getHeaderStyle());
        createCell(header, 1, "歸屬", styles.getHeaderStyle());
        createCell(header, 2, "代號", styles.getHeaderStyle());
        createCell(header, 3, "", styles.getHeaderStyle());
        createCell(header, 4, "險種", styles.getHeaderStyle());

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

            String categoryNumber = CategoryMapping.MAJOR_CATEGORY_NUMBER.getOrDefault(majorCat, "");
            int startRow = rowIdx;
            boolean firstInCategory = true;

            for (String code : codesInCategory) {
                Row row = sheet.createRow(rowIdx);
                // A: 類序號 (僅第一列)
                if (firstInCategory) {
                    createCell(row, 0, categoryNumber, styles.getCompanyStyle());
                    createCell(row, 1, majorCat, styles.getCompanyStyle());
                    firstInCategory = false;
                } else {
                    createCell(row, 0, "", styles.getCompanyStyle());
                    createCell(row, 1, "", styles.getCompanyStyle());
                }
                // C: 代號
                createCell(row, 2, code, styles.getCompanyStyle());
                // D: 子分組名 (僅在子分組第一個代號時填入)
                String subGroup = CategoryMapping.CODE_TO_SUB_GROUP.getOrDefault(code, "");
                createCell(row, 3, subGroup, styles.getCompanyStyle());
                // E: 險種全名
                createCell(row, 4, CategoryMapping.CODE_TO_FULL_NAME.getOrDefault(code, ""),
                        styles.getCompanyStyle());
                rowIdx++;
            }
        }

        sheet.setColumnWidth(0, 2000);
        sheet.setColumnWidth(1, 3000);
        sheet.setColumnWidth(2, 2500);
        sheet.setColumnWidth(3, 3000);
        sheet.setColumnWidth(4, 8000);
    }

    // ==================== Sheet 6: 排名 ====================

    private void writeRankingSheet(Workbook wb, ExcelStyleHelper styles,
                                   int year, int latestMonth,
                                   List<RankingCalculator.RankEntry> ranking) {
        Sheet sheet = wb.createSheet("排名");
        int rowIdx = 0;

        // 表頭 (對齊範例: 無標題行, 僅 公司 / 月份 兩欄)
        Row header = sheet.createRow(rowIdx++);
        // A 欄留空
        createCell(header, 1, "公司", styles.getHeaderStyle());
        String colHeader = latestMonth == 1
                ? "1月"
                : String.format("1-%d月", latestMonth);
        createCell(header, 2, colHeader, styles.getHeaderStyle());

        // 資料列
        for (RankingCalculator.RankEntry entry : ranking) {
            Row row = sheet.createRow(rowIdx++);
            createCell(row, 1, entry.getCompany().getName(), styles.getCompanyStyle());
            createCell(row, 2, entry.getTotalPremium(), styles.getNumberStyle());
        }

        sheet.setColumnWidth(0, 1500);
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
