package com.insurance.report.config;

import jakarta.validation.constraints.NotNull;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;
import org.springframework.validation.annotation.Validated;

import java.util.ArrayList;
import java.util.List;

/**
 * application.yml + insurance-mapping.yml 的 app.* 設定映射
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

    private InsuranceConfig insurance = new InsuranceConfig();

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

    public InsuranceConfig getInsurance() {
        return insurance;
    }

    public void setInsurance(InsuranceConfig insurance) {
        this.insurance = insurance;
    }

    // --- 內部類別 ---

    public static class ColumnsConfig {
        private List<String> hiddenCodes = new ArrayList<>();

        public List<String> getHiddenCodes() {
            return hiddenCodes;
        }

        public void setHiddenCodes(List<String> hiddenCodes) {
            this.hiddenCodes = hiddenCodes;
        }

        public boolean isIncludeOverseasReinsurance() {
            return !hiddenCodes.contains("9900");
        }
    }

    public enum CompanyOrder {
        BY_CODE_ASC,
        BY_CODE_DESC,
        BY_NAME
    }

    // --- 險種歸屬設定 ---

    public static class InsuranceConfig {
        private List<CodeConfig> codes = new ArrayList<>();
        private List<MajorCategoryConfig> categories = new ArrayList<>();

        public List<CodeConfig> getCodes() { return codes; }
        public void setCodes(List<CodeConfig> codes) { this.codes = codes; }
        public List<MajorCategoryConfig> getCategories() { return categories; }
        public void setCategories(List<MajorCategoryConfig> categories) { this.categories = categories; }
    }

    public static class CodeConfig {
        private String code;
        private String shortName;
        private String fullName;

        public String getCode() { return code; }
        public void setCode(String code) { this.code = code; }
        public String getShortName() { return shortName; }
        public void setShortName(String shortName) { this.shortName = shortName; }
        public String getFullName() { return fullName; }
        public void setFullName(String fullName) { this.fullName = fullName; }
    }

    public static class MajorCategoryConfig {
        private String name;
        private String number;
        private boolean overseas;
        private List<SubCategoryConfig> subCategories = new ArrayList<>();

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
        public String getNumber() { return number; }
        public void setNumber(String number) { this.number = number; }
        public boolean isOverseas() { return overseas; }
        public void setOverseas(boolean overseas) { this.overseas = overseas; }
        public List<SubCategoryConfig> getSubCategories() { return subCategories; }
        public void setSubCategories(List<SubCategoryConfig> subCategories) { this.subCategories = subCategories; }
    }

    public static class SubCategoryConfig {
        private String name;
        private String subGroup;
        private String headerGroup;
        private String headerLabel;
        private String shortHeader;
        private String subHeader;
        private List<String> codes = new ArrayList<>();

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
        public String getSubGroup() { return subGroup; }
        public void setSubGroup(String subGroup) { this.subGroup = subGroup; }
        public String getHeaderGroup() { return headerGroup; }
        public void setHeaderGroup(String headerGroup) { this.headerGroup = headerGroup; }
        public String getHeaderLabel() { return headerLabel; }
        public void setHeaderLabel(String headerLabel) { this.headerLabel = headerLabel; }
        public String getShortHeader() { return shortHeader; }
        public void setShortHeader(String shortHeader) { this.shortHeader = shortHeader; }
        public String getSubHeader() { return subHeader; }
        public void setSubHeader(String subHeader) { this.subHeader = subHeader; }
        public List<String> getCodes() { return codes; }
        public void setCodes(List<String> codes) { this.codes = codes; }
    }
}
