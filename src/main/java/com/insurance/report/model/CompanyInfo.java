package com.insurance.report.model;

import java.util.Objects;

/**
 * 公司基本資訊 (代號 + 名稱)
 */
public class CompanyInfo {

    private final String code;
    private final String name;

    public CompanyInfo(String code, String name) {
        this.code = code;
        this.name = name;
    }

    public String getCode() {
        return code;
    }

    public String getName() {
        return name;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CompanyInfo that = (CompanyInfo) o;
        return Objects.equals(code, that.code);
    }

    @Override
    public int hashCode() {
        return Objects.hash(code);
    }

    @Override
    public String toString() {
        return code + " " + name;
    }
}
