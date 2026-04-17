package com.insurance.report.model;

import com.insurance.report.config.AppConfig;
import com.insurance.report.config.AppConfig.*;
import jakarta.annotation.PostConstruct;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * 歸屬對照表 — 險種代號 → 九大類的分類規則
 * <p>
 * 從 config/insurance-mapping.yml 讀取設定，支援外部新增險種。
 */
@Component
public class CategoryMapping {

    private static final Logger log = LoggerFactory.getLogger(CategoryMapping.class);

    private final AppConfig config;

    // 以下欄位在 @PostConstruct 中從設定檔初始化
    private List<String> allInsuranceCodes;
    private Map<String, String> codeToShortName;
    private Map<String, String> codeToFullName;
    private List<SubCategory> subCategories;
    private SubCategory overseasReinsurance;
    private List<String> majorCategories;
    private Map<String, String> majorCategoryNumber;
    private Map<String, String> codeToSubGroup;

    public CategoryMapping(AppConfig config) {
        this.config = config;
    }

    @PostConstruct
    public void init() {
        InsuranceConfig ins = config.getInsurance();

        // 險種代號清單 (順序即明細表欄位順序)
        allInsuranceCodes = new ArrayList<>();
        codeToShortName = new LinkedHashMap<>();
        codeToFullName = new LinkedHashMap<>();
        for (CodeConfig cc : ins.getCodes()) {
            allInsuranceCodes.add(cc.getCode());
            codeToShortName.put(cc.getCode(), cc.getShortName());
            codeToFullName.put(cc.getCode(), cc.getFullName());
        }
        allInsuranceCodes = Collections.unmodifiableList(allInsuranceCodes);
        codeToShortName = Collections.unmodifiableMap(codeToShortName);
        codeToFullName = Collections.unmodifiableMap(codeToFullName);

        // 九大類 + 子分類
        majorCategories = new ArrayList<>();
        majorCategoryNumber = new LinkedHashMap<>();
        List<SubCategory> subList = new ArrayList<>();
        codeToSubGroup = new LinkedHashMap<>();
        SubCategory overseas = null;

        for (MajorCategoryConfig mc : ins.getCategories()) {
            majorCategories.add(mc.getName());
            majorCategoryNumber.put(mc.getName(), mc.getNumber());

            for (SubCategoryConfig sc : mc.getSubCategories()) {
                SubCategory sub = new SubCategory(sc.getName(), mc.getName(), sc.getCodes(),
                        sc.getHeaderGroup(), sc.getHeaderLabel(),
                        sc.getShortHeader(), sc.getSubHeader());
                if (mc.isOverseas()) {
                    overseas = sub;
                } else {
                    subList.add(sub);
                }
                // 子分組標籤 (歸屬表 D 欄)
                if (sc.getSubGroup() != null && !sc.getSubGroup().isEmpty()
                        && !sc.getCodes().isEmpty()) {
                    codeToSubGroup.put(sc.getCodes().get(0), sc.getSubGroup());
                }
            }
        }

        subCategories = Collections.unmodifiableList(subList);
        majorCategories = Collections.unmodifiableList(majorCategories);
        majorCategoryNumber = Collections.unmodifiableMap(majorCategoryNumber);
        codeToSubGroup = Collections.unmodifiableMap(codeToSubGroup);
        overseasReinsurance = overseas != null ? overseas
                : new SubCategory("國外分進", "國外分進", List.of("9900"));

        log.info("險種歸屬載入完成: {} 代號, {} 子分類, {} 大類",
                allInsuranceCodes.size(), subCategories.size(), majorCategories.size());
    }

    // --- Getters ---

    public List<String> getAllInsuranceCodes() {
        return allInsuranceCodes;
    }

    public Map<String, String> getCodeToShortName() {
        return codeToShortName;
    }

    public Map<String, String> getCodeToFullName() {
        return codeToFullName;
    }

    public List<SubCategory> getSubCategories() {
        return subCategories;
    }

    public SubCategory getOverseasReinsurance() {
        return overseasReinsurance;
    }

    public List<String> getMajorCategories() {
        return majorCategories;
    }

    public Map<String, String> getMajorCategoryNumber() {
        return majorCategoryNumber;
    }

    public Map<String, String> getCodeToSubGroup() {
        return codeToSubGroup;
    }

    /** 取得全部子分類 (含國外分進) */
    public List<SubCategory> getAllSubCategoriesWithOverseas() {
        List<SubCategory> list = new ArrayList<>(subCategories);
        list.add(overseasReinsurance);
        return list;
    }

    /**
     * 子分類定義
     */
    public static class SubCategory {
        private final String name;
        private final String majorCategory;
        private final List<String> insuranceCodes;
        private final String headerGroup;
        private final String headerLabel;
        private final String shortHeader;
        private final String subHeader;

        public SubCategory(String name, String majorCategory, List<String> insuranceCodes,
                           String headerGroup, String headerLabel,
                           String shortHeader, String subHeader) {
            this.name = name;
            this.majorCategory = majorCategory;
            this.insuranceCodes = Collections.unmodifiableList(new ArrayList<>(insuranceCodes));
            this.headerGroup = headerGroup;
            this.headerLabel = headerLabel;
            this.shortHeader = shortHeader;
            this.subHeader = subHeader;
        }

        public SubCategory(String name, String majorCategory, List<String> insuranceCodes) {
            this(name, majorCategory, insuranceCodes, null, null, null, null);
        }

        public String getName() { return name; }
        public String getMajorCategory() { return majorCategory; }
        public List<String> getInsuranceCodes() { return insuranceCodes; }
        public String getHeaderGroup() { return headerGroup; }
        public String getHeaderLabel() { return headerLabel; }
        public String getShortHeader() { return shortHeader; }
        public String getSubHeader() { return subHeader; }

        public long sumFrom(CompanyMonthData data) {
            return insuranceCodes.stream()
                    .mapToLong(data::getPremium)
                    .sum();
        }

        @Override
        public String toString() {
            return name + " (" + String.join(",", insuranceCodes) + ")";
        }
    }
}
