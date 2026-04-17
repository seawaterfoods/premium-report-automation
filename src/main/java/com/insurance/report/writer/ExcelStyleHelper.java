package com.insurance.report.writer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excel 格式工具 — 提供共用的儲存格樣式
 *
 * 色彩定義 (對照範例):
 *   背景: 淺青色 #CCFFFF (表頭, 左側欄位)
 *   數值字體: 藍色 #0000FF
 *   月份字體: 深藍 #000080, 粗體
 *   佔比標題: 紅色 #FF0000
 */
public class ExcelStyleHelper {

    private final Workbook workbook;

    // 色彩常數
    private static final byte[] CYAN_BG = {(byte) 0xCC, (byte) 0xFF, (byte) 0xFF};
    private static final byte[] BLUE_FONT = {0x00, 0x00, (byte) 0xFF};
    private static final byte[] DARK_BLUE = {0x00, 0x00, (byte) 0x80};
    private static final byte[] RED_FONT = {(byte) 0xFF, 0x00, 0x00};

    // === 保費統計表共用樣式 ===
    private CellStyle headerStyle;       // 表頭 (11pt 粗體, 青色底)
    private CellStyle subHeaderStyle;    // 子表頭 (11pt 非粗, 青色底)
    private CellStyle companyStyle;      // 代號/公司 (12pt, 青色底)
    private CellStyle monthStyle;        // 月份欄 (12pt 粗體深藍, 青色底)
    private CellStyle numberStyle;       // 數值 (10pt 藍字)
    private CellStyle subtotalStyle;     // 小計左側 (12pt, 青色底, 雙底線)
    private CellStyle subtotalNumberStyle; // 小計數值 (10pt 黑字, 雙底線)
    private CellStyle percentStyle;      // 百分比 (10pt 黑字)
    private CellStyle titleStyle;        // 標題 (18pt 粗體)

    // === 比較分析表專用樣式 ===
    private CellStyle cmpTitleStyle;     // 比較標題 (14pt 粗體)
    private CellStyle cmpPeriodStyle;    // 月份/期間 (12pt 粗體, 青色底)
    private CellStyle cmpHeaderStyle;    // 比較表頭 (11pt 粗體, 青色底)
    private CellStyle cmpSubHeaderStyle; // 比較子表頭 (10pt, 青色底)
    private CellStyle cmpLabelStyle;     // 行標籤 (10pt)
    private CellStyle cmpNumberStyle;    // 比較數值 (10pt)
    private CellStyle cmpBlueNumberStyle;// 去年值/藍字 (10pt 藍字)
    private CellStyle cmpPercentStyle;   // 百分比 (10pt)
    private CellStyle cmpRedLabelStyle;  // 佔比標題 (10pt 紅字)
    private CellStyle cmpBoldPercentStyle; // 成長率 (10pt 粗體)

    public ExcelStyleHelper(Workbook workbook) {
        this.workbook = workbook;
        initStyles();
    }

    private void initStyles() {
        XSSFWorkbook xwb = (XSSFWorkbook) workbook;
        XSSFColor cyanColor = new XSSFColor(CYAN_BG, null);
        DataFormat df = workbook.createDataFormat();

        // --- 字體定義 ---
        XSSFFont titleFont = xwb.createFont();
        titleFont.setBold(true);
        titleFont.setFontName("微軟正黑體");
        titleFont.setFontHeightInPoints((short) 18);

        XSSFFont headerFont = xwb.createFont();
        headerFont.setBold(true);
        headerFont.setFontName("微軟正黑體");
        headerFont.setFontHeightInPoints((short) 11);

        XSSFFont normalFont = xwb.createFont();
        normalFont.setFontName("微軟正黑體");
        normalFont.setFontHeightInPoints((short) 11);

        XSSFFont companyFont = xwb.createFont();
        companyFont.setFontName("微軟正黑體");
        companyFont.setFontHeightInPoints((short) 12);

        XSSFFont monthFont = xwb.createFont();
        monthFont.setBold(true);
        monthFont.setFontName("微軟正黑體");
        monthFont.setFontHeightInPoints((short) 12);
        monthFont.setColor(new XSSFColor(DARK_BLUE, null));

        XSSFFont dataFont = xwb.createFont();
        dataFont.setFontName("微軟正黑體");
        dataFont.setFontHeightInPoints((short) 10);
        dataFont.setColor(new XSSFColor(BLUE_FONT, null));

        XSSFFont normalDataFont = xwb.createFont();
        normalDataFont.setFontName("微軟正黑體");
        normalDataFont.setFontHeightInPoints((short) 10);

        XSSFFont redFont = xwb.createFont();
        redFont.setFontName("微軟正黑體");
        redFont.setFontHeightInPoints((short) 10);
        redFont.setColor(new XSSFColor(RED_FONT, null));

        XSSFFont boldDataFont = xwb.createFont();
        boldDataFont.setBold(true);
        boldDataFont.setFontName("微軟正黑體");
        boldDataFont.setFontHeightInPoints((short) 10);

        // --- 保費統計表樣式 ---

        // 標題
        titleStyle = xwb.createCellStyle();
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 表頭 (粗體 + 青色底 + 框線)
        headerStyle = xwb.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);
        applyCyanBg((XSSFCellStyle) headerStyle, cyanColor);
        applyThinBorders(headerStyle);

        // 子表頭
        subHeaderStyle = xwb.createCellStyle();
        subHeaderStyle.setFont(normalFont);
        subHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        subHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        subHeaderStyle.setWrapText(true);
        applyCyanBg((XSSFCellStyle) subHeaderStyle, cyanColor);
        applyThinBorders(subHeaderStyle);

        // 代號/公司名
        companyStyle = xwb.createCellStyle();
        companyStyle.setFont(companyFont);
        companyStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) companyStyle, cyanColor);
        applyThinBorders(companyStyle);

        // 月份欄
        monthStyle = xwb.createCellStyle();
        monthStyle.setFont(monthFont);
        monthStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        monthStyle.setAlignment(HorizontalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) monthStyle, cyanColor);
        applyThinBorders(monthStyle);

        // 數值 (藍字)
        numberStyle = xwb.createCellStyle();
        numberStyle.setFont(dataFont);
        numberStyle.setDataFormat(df.getFormat("#,##0"));
        numberStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(numberStyle);

        // 小計左側
        subtotalStyle = xwb.createCellStyle();
        subtotalStyle.setFont(companyFont);
        subtotalStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) subtotalStyle, cyanColor);
        applyThinBorders(subtotalStyle);
        subtotalStyle.setBorderBottom(BorderStyle.DOUBLE);

        // 小計數值 (黑字, 雙底線)
        subtotalNumberStyle = xwb.createCellStyle();
        subtotalNumberStyle.setFont(normalDataFont);
        subtotalNumberStyle.setDataFormat(df.getFormat("#,##0"));
        subtotalNumberStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(subtotalNumberStyle);
        subtotalNumberStyle.setBorderBottom(BorderStyle.DOUBLE);

        // 百分比
        percentStyle = xwb.createCellStyle();
        percentStyle.setFont(normalDataFont);
        percentStyle.setDataFormat(df.getFormat("0.00%"));
        percentStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(percentStyle);

        // --- 比較分析表樣式 ---

        XSSFFont cmpTitleFont = xwb.createFont();
        cmpTitleFont.setBold(true);
        cmpTitleFont.setFontName("微軟正黑體");
        cmpTitleFont.setFontHeightInPoints((short) 14);

        XSSFFont cmpPeriodFont = xwb.createFont();
        cmpPeriodFont.setBold(true);
        cmpPeriodFont.setFontName("微軟正黑體");
        cmpPeriodFont.setFontHeightInPoints((short) 12);

        cmpTitleStyle = xwb.createCellStyle();
        cmpTitleStyle.setFont(cmpTitleFont);
        cmpTitleStyle.setAlignment(HorizontalAlignment.CENTER);
        cmpTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cmpPeriodStyle = xwb.createCellStyle();
        cmpPeriodStyle.setFont(cmpPeriodFont);
        cmpPeriodStyle.setAlignment(HorizontalAlignment.LEFT);
        applyCyanBg((XSSFCellStyle) cmpPeriodStyle, cyanColor);

        cmpHeaderStyle = xwb.createCellStyle();
        cmpHeaderStyle.setFont(headerFont);
        cmpHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        cmpHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cmpHeaderStyle.setWrapText(true);
        applyCyanBg((XSSFCellStyle) cmpHeaderStyle, cyanColor);
        applyThinBorders(cmpHeaderStyle);

        cmpSubHeaderStyle = xwb.createCellStyle();
        cmpSubHeaderStyle.setFont(normalDataFont);
        cmpSubHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        cmpSubHeaderStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyCyanBg((XSSFCellStyle) cmpSubHeaderStyle, cyanColor);
        applyThinBorders(cmpSubHeaderStyle);

        cmpLabelStyle = xwb.createCellStyle();
        cmpLabelStyle.setFont(normalDataFont);
        cmpLabelStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(cmpLabelStyle);

        cmpNumberStyle = xwb.createCellStyle();
        cmpNumberStyle.setFont(normalDataFont);
        cmpNumberStyle.setDataFormat(df.getFormat("#,##0"));
        cmpNumberStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(cmpNumberStyle);

        cmpBlueNumberStyle = xwb.createCellStyle();
        cmpBlueNumberStyle.setFont(dataFont);
        cmpBlueNumberStyle.setDataFormat(df.getFormat("#,##0"));
        cmpBlueNumberStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(cmpBlueNumberStyle);

        cmpPercentStyle = xwb.createCellStyle();
        cmpPercentStyle.setFont(normalDataFont);
        cmpPercentStyle.setDataFormat(df.getFormat("0.00%"));
        cmpPercentStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(cmpPercentStyle);

        cmpRedLabelStyle = xwb.createCellStyle();
        cmpRedLabelStyle.setFont(redFont);
        cmpRedLabelStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorders(cmpRedLabelStyle);

        cmpBoldPercentStyle = xwb.createCellStyle();
        cmpBoldPercentStyle.setFont(boldDataFont);
        cmpBoldPercentStyle.setDataFormat(df.getFormat("0.00%"));
        cmpBoldPercentStyle.setAlignment(HorizontalAlignment.RIGHT);
        applyThinBorders(cmpBoldPercentStyle);
    }

    private void applyCyanBg(XSSFCellStyle style, XSSFColor color) {
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }

    private void applyThinBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    // --- 保費統計表 getters ---
    public CellStyle getHeaderStyle() { return headerStyle; }
    public CellStyle getSubHeaderStyle() { return subHeaderStyle; }
    public CellStyle getCompanyStyle() { return companyStyle; }
    public CellStyle getMonthStyle() { return monthStyle; }
    public CellStyle getNumberStyle() { return numberStyle; }
    public CellStyle getSubtotalStyle() { return subtotalStyle; }
    public CellStyle getSubtotalNumberStyle() { return subtotalNumberStyle; }
    public CellStyle getPercentStyle() { return percentStyle; }
    public CellStyle getTitleStyle() { return titleStyle; }

    // --- 比較分析表 getters ---
    public CellStyle getCmpTitleStyle() { return cmpTitleStyle; }
    public CellStyle getCmpPeriodStyle() { return cmpPeriodStyle; }
    public CellStyle getCmpHeaderStyle() { return cmpHeaderStyle; }
    public CellStyle getCmpSubHeaderStyle() { return cmpSubHeaderStyle; }
    public CellStyle getCmpLabelStyle() { return cmpLabelStyle; }
    public CellStyle getCmpNumberStyle() { return cmpNumberStyle; }
    public CellStyle getCmpBlueNumberStyle() { return cmpBlueNumberStyle; }
    public CellStyle getCmpPercentStyle() { return cmpPercentStyle; }
    public CellStyle getCmpRedLabelStyle() { return cmpRedLabelStyle; }
    public CellStyle getCmpBoldPercentStyle() { return cmpBoldPercentStyle; }

    // --- 工具方法 ---

    /** 建立儲存格並設定文字值 + 樣式 */
    public static Cell createCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        if (value != null) cell.setCellValue(value);
        if (style != null) cell.setCellStyle(style);
        return cell;
    }

    /** 建立儲存格並設定數字值 + 樣式 */
    public static Cell createCell(Row row, int col, long value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value);
        if (style != null) cell.setCellStyle(style);
        return cell;
    }

    /** 建立儲存格並設定浮點數值 + 樣式 */
    public static Cell createCell(Row row, int col, double value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value);
        if (style != null) cell.setCellStyle(style);
        return cell;
    }

    /** 建立公式儲存格 */
    public static Cell createFormulaCell(Row row, int col, String formula, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellFormula(formula);
        if (style != null) cell.setCellStyle(style);
        return cell;
    }

    /**
     * 0-based 欄位索引 → Excel 欄位字母 (A, B, ..., Z, AA, AB, ...)
     */
    public static String colLetter(int colIndex) {
        StringBuilder sb = new StringBuilder();
        int idx = colIndex;
        while (idx >= 0) {
            sb.insert(0, (char) ('A' + idx % 26));
            idx = idx / 26 - 1;
        }
        return sb.toString();
    }

    /**
     * 產生儲存格參考 (e.g., "D6")
     * @param col 0-based 欄位索引
     * @param row 0-based 列索引
     */
    public static String cellRef(int col, int row) {
        return colLetter(col) + (row + 1);
    }

    /** 合併儲存格 */
    public static void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        if (firstRow == lastRow && firstCol == lastCol) return;
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    /** 為區域填充邊框樣式 (使用 companyStyle 作為預設) */
    public void fillBorders(Sheet sheet, int startRow, int endRow, int startCol, int endCol) {
        fillBorders(sheet, startRow, endRow, startCol, endCol, companyStyle);
    }

    /** 為區域填充邊框樣式 (使用指定樣式作為預設) */
    public void fillBorders(Sheet sheet, int startRow, int endRow, int startCol, int endCol, CellStyle defaultStyle) {
        for (int r = startRow; r <= endRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) row = sheet.createRow(r);
            for (int c = startCol; c <= endCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) {
                    cell = row.createCell(c);
                    cell.setCellStyle(defaultStyle);
                }
            }
        }
    }
}
