package com.insurance.report.calculator;

import com.insurance.report.config.AppConfig;
import com.insurance.report.model.CategoryMapping;
import com.insurance.report.model.CategoryMapping.SubCategory;
import com.insurance.report.model.CompanyMonthData;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * 九大類（15 子分類）彙總計算
 * <p>
 * 將 33 險種按歸屬規則加總為 15 個子分類欄位，
 * 用於「總」表 (Sheet 3/4) 和比較表。
 */
@Component
public class CategoryCalculator {

    private final AppConfig config;

    public CategoryCalculator(AppConfig config) {
        this.config = config;
    }

    /**
     * 將公司月資料的 33 險種轉換為子分類金額
     *
     * @return 子分類名稱 → 金額 (LinkedHashMap 保持順序)
     */
    public Map<String, Long> calculateSubCategories(CompanyMonthData data) {
        Map<String, Long> result = new LinkedHashMap<>();

        for (SubCategory sub : CategoryMapping.SUB_CATEGORIES) {
            result.put(sub.getName(), sub.sumFrom(data));
        }

        // 國外分進 (依設定決定是否包含)
        if (config.getColumns().isIncludeOverseasReinsurance()) {
            result.put(CategoryMapping.OVERSEAS_REINSURANCE.getName(),
                    CategoryMapping.OVERSEAS_REINSURANCE.sumFrom(data));
        }

        return result;
    }

    /**
     * 將公司月資料的 33 險種轉換為全部 16 子分類金額 (含國外分進)
     */
    public Map<String, Long> calculateAllSubCategories(CompanyMonthData data) {
        Map<String, Long> result = new LinkedHashMap<>();
        for (SubCategory sub : CategoryMapping.SUB_CATEGORIES) {
            result.put(sub.getName(), sub.sumFrom(data));
        }
        result.put(CategoryMapping.OVERSEAS_REINSURANCE.getName(),
                CategoryMapping.OVERSEAS_REINSURANCE.sumFrom(data));
        return result;
    }

    /**
     * 計算合計 (D~R 或 D~S 加總，取決於是否包含國外分進)
     */
    public long calculateTotal(Map<String, Long> subCategoryAmounts) {
        return subCategoryAmounts.values().stream().mapToLong(Long::longValue).sum();
    }

    /**
     * 計算輸出用合計 (依設定包含/排除國外分進)
     */
    public long calculateOutputTotal(Map<String, Long> allSubCategoryAmounts) {
        long total = 0;
        for (SubCategory sub : getOutputSubCategories()) {
            total += allSubCategoryAmounts.getOrDefault(sub.getName(), 0L);
        }
        return total;
    }

    /**
     * 取得要輸出的子分類清單 (含/不含國外分進)
     */
    public List<SubCategory> getOutputSubCategories() {
        List<SubCategory> list = new ArrayList<>(CategoryMapping.SUB_CATEGORIES);
        if (config.getColumns().isIncludeOverseasReinsurance()) {
            list.add(CategoryMapping.OVERSEAS_REINSURANCE);
        }
        return list;
    }
}
