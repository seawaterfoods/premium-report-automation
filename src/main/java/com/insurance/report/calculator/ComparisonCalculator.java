package com.insurance.report.calculator;

import com.insurance.report.model.GrowthRateResult;
import com.insurance.report.util.GrowthRateUtil;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * 同期比較計算
 * <p>
 * 計算今年 vs 去年的增減金額與成長率。
 * 成長率公式: (今年/去年) - 1
 * 去年=0 → 分母=1, 記錄警告(缺資料)
 * 去年&lt;0 → 分母=1, 記錄警告(負數資料)
 */
@Component
public class ComparisonCalculator {

    /**
     * 比較結果
     */
    public static class ComparisonRow {
        private final Map<String, Long> currentValues;
        private final Map<String, Long> priorValues;
        private final Map<String, Long> differences;
        private final Map<String, GrowthRateResult> growthRates;
        private final Map<String, Double> proportions;

        public ComparisonRow(Map<String, Long> currentValues, Map<String, Long> priorValues,
                             Map<String, Long> differences, Map<String, GrowthRateResult> growthRates,
                             Map<String, Double> proportions) {
            this.currentValues = currentValues;
            this.priorValues = priorValues;
            this.differences = differences;
            this.growthRates = growthRates;
            this.proportions = proportions;
        }

        public Map<String, Long> getCurrentValues() { return currentValues; }
        public Map<String, Long> getPriorValues() { return priorValues; }
        public Map<String, Long> getDifferences() { return differences; }
        public Map<String, GrowthRateResult> getGrowthRates() { return growthRates; }
        public Map<String, Double> getProportions() { return proportions; }
    }

    /**
     * 計算同期比較
     *
     * @param currentSubtotals 今年子分類小計 (類別名→金額)
     * @param priorSubtotals   去年子分類小計 (類別名→金額)
     * @param categories       要比較的類別名稱清單
     * @param periodLabel      期間描述 (用於日誌, e.g., "3月" 或 "1-3月累計")
     */
    public ComparisonRow calculate(
            Map<String, Long> currentSubtotals,
            Map<String, Long> priorSubtotals,
            List<String> categories,
            String periodLabel) {

        Map<String, Long> currentValues = new LinkedHashMap<>();
        Map<String, Long> priorValues = new LinkedHashMap<>();
        Map<String, Long> differences = new LinkedHashMap<>();
        Map<String, GrowthRateResult> growthRates = new LinkedHashMap<>();
        Map<String, Double> proportions = new LinkedHashMap<>();

        // 今年合計 (用於計算佔比)
        long currentTotal = 0;
        for (String cat : categories) {
            long cv = currentSubtotals.getOrDefault(cat, 0L);
            currentTotal += cv;
        }

        for (String cat : categories) {
            long cv = currentSubtotals.getOrDefault(cat, 0L);
            long pv = priorSubtotals.getOrDefault(cat, 0L);

            currentValues.put(cat, cv);
            priorValues.put(cat, pv);
            differences.put(cat, cv - pv);

            String context = cat + "/" + periodLabel;
            growthRates.put(cat, GrowthRateUtil.calculate(cv, pv, context));

            // 佔比 = 該類/合計
            double proportion = currentTotal == 0 ? 0.0 : (double) cv / currentTotal;
            proportions.put(cat, proportion);
        }

        // 合計列
        long totalCurrent = currentValues.values().stream().mapToLong(Long::longValue).sum();
        long totalPrior = priorValues.values().stream().mapToLong(Long::longValue).sum();
        currentValues.put("合計", totalCurrent);
        priorValues.put("合計", totalPrior);
        differences.put("合計", totalCurrent - totalPrior);
        growthRates.put("合計", GrowthRateUtil.calculate(totalCurrent, totalPrior, "合計/" + periodLabel));
        proportions.put("合計", 1.0);

        return new ComparisonRow(currentValues, priorValues, differences, growthRates, proportions);
    }
}
