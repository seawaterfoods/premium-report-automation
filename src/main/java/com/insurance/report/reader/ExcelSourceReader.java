package com.insurance.report.reader;

import com.insurance.report.model.CompanyMonthData;
import com.insurance.report.model.SourceFileInfo;
import com.insurance.report.util.InsuranceCodeUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

/**
 * 讀取來源 Excel 檔案 (保險收入統計表.xlsx)
 * <p>
 * Sheet1「保險收入統計表」:
 *   C2 = 年月(int), B4 = 公司代碼(str), C4 = 公司名稱(str)
 *   A6:A38 = 險種代號(33筆), C6:C38 = 金額(int)
 *   C39 = 合計
 * <p>
 * Sheet2「歸屬」: 分類對照表 (可選)
 */
@Component
public class ExcelSourceReader {

    private static final Logger log = LoggerFactory.getLogger(ExcelSourceReader.class);

    // Sheet1 固定位置
    private static final int YEAR_MONTH_ROW = 1;    // C2 (0-based: row 1)
    private static final int YEAR_MONTH_COL = 2;    // col C
    private static final int COMPANY_CODE_ROW = 3;  // B4 (0-based: row 3)
    private static final int COMPANY_CODE_COL = 1;  // col B
    private static final int COMPANY_NAME_COL = 2;  // col C
    private static final int DATA_START_ROW = 5;     // A6 (0-based: row 5)
    private static final int DATA_END_ROW = 37;      // A38 (0-based: row 37), 33 rows
    private static final int CODE_COL = 0;           // col A
    private static final int AMOUNT_COL = 2;         // col C

    /**
     * 讀取來源檔案，回傳公司月資料
     */
    public CompanyMonthData read(SourceFileInfo fileInfo) throws IOException {
        log.debug("讀取來源檔: {}", fileInfo.getFilePath().getFileName());

        try (InputStream is = Files.newInputStream(fileInfo.getFilePath());
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheet("保險收入統計表");
            if (sheet == null) {
                sheet = workbook.getSheetAt(0);
                log.warn("找不到「保險收入統計表」分頁，使用第一個分頁: {}",
                        sheet.getSheetName());
            }

            // 讀取公司資訊
            String companyCode = readCellAsString(sheet, COMPANY_CODE_ROW, COMPANY_CODE_COL);
            String companyName = readCellAsString(sheet, COMPANY_CODE_ROW, COMPANY_NAME_COL);

            // 公司代號左補零至 2 位
            if (companyCode != null && companyCode.length() == 1) {
                companyCode = "0" + companyCode;
            }

            CompanyMonthData data = new CompanyMonthData(
                    companyCode, companyName, fileInfo.getYear(), fileInfo.getMonth());

            // 讀取 33 筆險種金額
            List<String> warnings = new ArrayList<>();
            for (int row = DATA_START_ROW; row <= DATA_END_ROW; row++) {
                Row r = sheet.getRow(row);
                if (r == null) continue;

                // 險種代號
                Cell codeCell = r.getCell(CODE_COL);
                if (codeCell == null) continue;

                String rawCode = readCellValue(codeCell);
                String code = InsuranceCodeUtil.normalize(rawCode);
                if (code == null || !InsuranceCodeUtil.isValid(code)) {
                    warnings.add(String.format("Row %d: 無效的險種代號 '%s'", row + 1, rawCode));
                    continue;
                }

                // 金額
                long amount = readCellAsLong(sheet, row, AMOUNT_COL);
                data.putPremium(code, amount);
            }

            data.recalculateTotal();

            if (!warnings.isEmpty()) {
                log.warn("來源檔 {} 讀取警告:\n  {}", fileInfo.getFilePath().getFileName(),
                        String.join("\n  ", warnings));
            }

            // 驗證歸屬分頁是否存在
            Sheet guishuSheet = workbook.getSheet("歸屬");
            if (guishuSheet == null) {
                log.error("公司別{}({})，缺少歸屬分頁", companyCode, companyName);
            }

            log.debug("讀取完成: {}", data);
            return data;
        }
    }

    /**
     * 僅讀取檔案內的公司代號 (用於檔名比對檢核)
     */
    public String readCompanyCode(SourceFileInfo fileInfo) throws IOException {
        try (InputStream is = Files.newInputStream(fileInfo.getFilePath());
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheet("保險收入統計表");
            if (sheet == null) {
                sheet = workbook.getSheetAt(0);
            }

            String code = readCellAsString(sheet, COMPANY_CODE_ROW, COMPANY_CODE_COL);
            if (code != null && code.length() == 1) {
                code = "0" + code;
            }
            return code;
        }
    }

    private String readCellAsString(Sheet sheet, int row, int col) {
        Row r = sheet.getRow(row);
        if (r == null) return null;
        Cell cell = r.getCell(col);
        if (cell == null) return null;
        return readCellValue(cell);
    }

    private long readCellAsLong(Sheet sheet, int row, int col) {
        Row r = sheet.getRow(row);
        if (r == null) return 0;
        Cell cell = r.getCell(col);
        if (cell == null) return 0;

        if (cell.getCellType() == CellType.NUMERIC) {
            return (long) cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return Long.parseLong(cell.getStringCellValue().trim().replace(",", ""));
            } catch (NumberFormatException e) {
                return 0;
            }
        }
        return 0;
    }

    private String readCellValue(Cell cell) {
        if (cell == null) return null;
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                double val = cell.getNumericCellValue();
                if (val == Math.floor(val) && !Double.isInfinite(val)) {
                    return String.valueOf((long) val);
                }
                return String.valueOf(val);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return String.valueOf((long) cell.getNumericCellValue());
                } catch (Exception e) {
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception e2) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return null;
        }
    }
}
