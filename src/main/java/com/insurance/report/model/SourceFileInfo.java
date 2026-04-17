package com.insurance.report.model;

import java.nio.file.Path;

/**
 * 已掃描到的來源檔案資訊
 */
public class SourceFileInfo {

    private final Path filePath;
    private final String companyCode;
    /** 民國年 */
    private final int year;
    /** 月份 (1~12) */
    private final int month;

    public SourceFileInfo(Path filePath, String companyCode, int year, int month) {
        this.filePath = filePath;
        this.companyCode = companyCode;
        this.year = year;
        this.month = month;
    }

    public Path getFilePath() {
        return filePath;
    }

    public String getCompanyCode() {
        return companyCode;
    }

    public int getYear() {
        return year;
    }

    public int getMonth() {
        return month;
    }

    /** 民國年月 4 碼 (e.g., 11503) */
    public int getYearMonth() {
        return year * 100 + month;
    }

    @Override
    public String toString() {
        return String.format("SourceFile{%s, company=%s, %d/%02d}", filePath.getFileName(), companyCode, year, month);
    }
}
