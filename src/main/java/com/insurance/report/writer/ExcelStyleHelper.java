package com.insurance.report.writer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excel 格式工具 — 提供共用的儲存格樣式
 */
public class ExcelStyleHelper {

    private final Workbook workbook;

    // 快取常用樣式
    private CellStyle headerStyle;
    private CellStyle subHeaderStyle;
    private CellStyle companyStyle;
    private CellStyle numberStyle;
    private CellStyle subtotalStyle;
    private CellStyle subtotalNumberStyle;
    private CellStyle percentStyle;
    private CellStyle titleStyle;

    public ExcelStyleHelper(Workbook workbook) {
        this.workbook = workbook;
        initStyles();
    }

    private void initStyles() {
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        boldFont.setFontName("微軟正黑體");
        boldFont.setFontHeightInPoints((short) 10);

        Font normalFont = workbook.createFont();
        normalFont.setFontName("微軟正黑體");
        normalFont.setFontHeightInPoints((short) 10);

        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontName("微軟正黑體");
        titleFont.setFontHeightInPoints((short) 12);

        // 標題
        titleStyle = workbook.createCellStyle();
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 表頭
        headerStyle = workbook.createCellStyle();
        headerStyle.setFont(boldFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setWrapText(true);

        // 子表頭
        subHeaderStyle = workbook.createCellStyle();
        subHeaderStyle.cloneStyleFrom(headerStyle);
        subHeaderStyle.setFont(normalFont);

        // 公司名稱
        companyStyle = workbook.createCellStyle();
        companyStyle.setFont(normalFont);
        companyStyle.setBorderTop(BorderStyle.THIN);
        companyStyle.setBorderBottom(BorderStyle.THIN);
        companyStyle.setBorderLeft(BorderStyle.THIN);
        companyStyle.setBorderRight(BorderStyle.THIN);
        companyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // 數字格式
        DataFormat df = workbook.createDataFormat();
        numberStyle = workbook.createCellStyle();
        numberStyle.setFont(normalFont);
        numberStyle.setDataFormat(df.getFormat("#,##0"));
        numberStyle.setBorderTop(BorderStyle.THIN);
        numberStyle.setBorderBottom(BorderStyle.THIN);
        numberStyle.setBorderLeft(BorderStyle.THIN);
        numberStyle.setBorderRight(BorderStyle.THIN);
        numberStyle.setAlignment(HorizontalAlignment.RIGHT);

        // 小計列
        subtotalStyle = workbook.createCellStyle();
        subtotalStyle.setFont(boldFont);
        subtotalStyle.setBorderTop(BorderStyle.THIN);
        subtotalStyle.setBorderBottom(BorderStyle.DOUBLE);
        subtotalStyle.setBorderLeft(BorderStyle.THIN);
        subtotalStyle.setBorderRight(BorderStyle.THIN);
        subtotalStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        subtotalNumberStyle = workbook.createCellStyle();
        subtotalNumberStyle.setFont(boldFont);
        subtotalNumberStyle.setDataFormat(df.getFormat("#,##0"));
        subtotalNumberStyle.setBorderTop(BorderStyle.THIN);
        subtotalNumberStyle.setBorderBottom(BorderStyle.DOUBLE);
        subtotalNumberStyle.setBorderLeft(BorderStyle.THIN);
        subtotalNumberStyle.setBorderRight(BorderStyle.THIN);
        subtotalNumberStyle.setAlignment(HorizontalAlignment.RIGHT);

        // 百分比格式
        percentStyle = workbook.createCellStyle();
        percentStyle.setFont(normalFont);
        percentStyle.setDataFormat(df.getFormat("0.00%"));
        percentStyle.setBorderTop(BorderStyle.THIN);
        percentStyle.setBorderBottom(BorderStyle.THIN);
        percentStyle.setBorderLeft(BorderStyle.THIN);
        percentStyle.setBorderRight(BorderStyle.THIN);
        percentStyle.setAlignment(HorizontalAlignment.RIGHT);
    }

    public CellStyle getHeaderStyle() { return headerStyle; }
    public CellStyle getSubHeaderStyle() { return subHeaderStyle; }
    public CellStyle getCompanyStyle() { return companyStyle; }
    public CellStyle getNumberStyle() { return numberStyle; }
    public CellStyle getSubtotalStyle() { return subtotalStyle; }
    public CellStyle getSubtotalNumberStyle() { return subtotalNumberStyle; }
    public CellStyle getPercentStyle() { return percentStyle; }
    public CellStyle getTitleStyle() { return titleStyle; }

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

    /** 合併儲存格 */
    public static void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        if (firstRow == lastRow && firstCol == lastCol) return;
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    /** 為區域填充邊框樣式 */
    public void fillBorders(Sheet sheet, int startRow, int endRow, int startCol, int endCol) {
        for (int r = startRow; r <= endRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) row = sheet.createRow(r);
            for (int c = startCol; c <= endCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) {
                    cell = row.createCell(c);
                    cell.setCellStyle(companyStyle);
                }
            }
        }
    }
}
