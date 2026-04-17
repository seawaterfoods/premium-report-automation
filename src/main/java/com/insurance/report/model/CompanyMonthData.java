package com.insurance.report.model;

import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Objects;

/**
 * 單一公司、單一月份的保險收入資料
 */
public class CompanyMonthData {

    private String companyCode;
    private String companyName;
    /** 民國年 (e.g., 115) */
    private int year;
    /** 月份 (1~12) */
    private int month;
    /** 險種代號 → 保險收入金額 (代號統一為文字) */
    private final Map<String, Long> premiums = new LinkedHashMap<>();
    /** 合計 */
    private long total;

    public CompanyMonthData() {
    }

    public CompanyMonthData(String companyCode, String companyName, int year, int month) {
        this.companyCode = companyCode;
        this.companyName = companyName;
        this.year = year;
        this.month = month;
    }

    public void putPremium(String insuranceCode, long amount) {
        premiums.put(insuranceCode, amount);
    }

    public long getPremium(String insuranceCode) {
        return premiums.getOrDefault(insuranceCode, 0L);
    }

    // --- getters & setters ---

    public String getCompanyCode() {
        return companyCode;
    }

    public void setCompanyCode(String companyCode) {
        this.companyCode = companyCode;
    }

    public String getCompanyName() {
        return companyName;
    }

    public void setCompanyName(String companyName) {
        this.companyName = companyName;
    }

    public int getYear() {
        return year;
    }

    public void setYear(int year) {
        this.year = year;
    }

    public int getMonth() {
        return month;
    }

    public void setMonth(int month) {
        this.month = month;
    }

    public Map<String, Long> getPremiums() {
        return premiums;
    }

    public long getTotal() {
        return total;
    }

    public void setTotal(long total) {
        this.total = total;
    }

    /** 重新計算合計 (所有險種金額加總) */
    public void recalculateTotal() {
        this.total = premiums.values().stream().mapToLong(Long::longValue).sum();
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CompanyMonthData that = (CompanyMonthData) o;
        return year == that.year && month == that.month
                && Objects.equals(companyCode, that.companyCode);
    }

    @Override
    public int hashCode() {
        return Objects.hash(companyCode, year, month);
    }

    @Override
    public String toString() {
        return String.format("CompanyMonthData{code=%s, name=%s, %d/%02d, types=%d, total=%d}",
                companyCode, companyName, year, month, premiums.size(), total);
    }
}
