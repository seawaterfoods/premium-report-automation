package com.insurance.report.config;

import jakarta.validation.constraints.NotNull;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;
import org.springframework.validation.annotation.Validated;

import java.util.ArrayList;
import java.util.List;

/**
 * application.yml 的 app.* 設定映射
 */
@Component
@ConfigurationProperties(prefix = "app")
@Validated
public class AppConfig {

    @NotNull
    private String importDir = "./import";

    @NotNull
    private String outputDir = "./output";

    private Integer processYear;

    private ColumnsConfig columns = new ColumnsConfig();

    private CompanyOrder companyOrder = CompanyOrder.BY_CODE_ASC;

    // --- getters & setters ---

    public String getImportDir() {
        return importDir;
    }

    public void setImportDir(String importDir) {
        this.importDir = importDir;
    }

    public String getOutputDir() {
        return outputDir;
    }

    public void setOutputDir(String outputDir) {
        this.outputDir = outputDir;
    }

    public Integer getProcessYear() {
        return processYear;
    }

    public void setProcessYear(Integer processYear) {
        this.processYear = processYear;
    }

    public ColumnsConfig getColumns() {
        return columns;
    }

    public void setColumns(ColumnsConfig columns) {
        this.columns = columns;
    }

    public CompanyOrder getCompanyOrder() {
        return companyOrder;
    }

    public void setCompanyOrder(CompanyOrder companyOrder) {
        this.companyOrder = companyOrder;
    }

    // --- 內部類別 ---

    public static class ColumnsConfig {
        private boolean includeOverseasReinsurance = false;
        private List<String> hiddenCodes = new ArrayList<>();

        public boolean isIncludeOverseasReinsurance() {
            return includeOverseasReinsurance;
        }

        public void setIncludeOverseasReinsurance(boolean includeOverseasReinsurance) {
            this.includeOverseasReinsurance = includeOverseasReinsurance;
        }

        public List<String> getHiddenCodes() {
            return hiddenCodes;
        }

        public void setHiddenCodes(List<String> hiddenCodes) {
            this.hiddenCodes = hiddenCodes;
        }
    }

    public enum CompanyOrder {
        BY_CODE_ASC,
        BY_CODE_DESC,
        BY_NAME
    }
}
