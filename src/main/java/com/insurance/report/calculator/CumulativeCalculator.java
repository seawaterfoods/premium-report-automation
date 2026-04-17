package com.insurance.report.calculator;

import com.insurance.report.model.CategoryMapping;
import com.insurance.report.model.CompanyMonthData;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * 累計彙整計算
 * <p>
 * 累計僅產出至最新月份 (從 1 月起算)。
 * 例如只有 1~3 月資料，產出：1, 1-2, 1-3
 */
@Component
public class CumulativeCalculator {

    /**
     * 計算累計資料
     *
     * @param monthlyData  按月份整理的單月資料 (1~12)
     * @param companyList  公司清單
     * @param latestMonth  最新月份
     * @return 累計期間標籤 → List&lt;CompanyMonthData&gt;
     *         key 為累計期間 (1, 2, 3, ...)，表示 "1~key月" 的累計
     */
    public Map<Integer, List<CompanyMonthData>> calculate(
            Map<Integer, List<CompanyMonthData>> monthlyData,
            List<String> companyList,
            int latestMonth) {

        Map<Integer, List<CompanyMonthData>> result = new LinkedHashMap<>();

        for (int endMonth = 1; endMonth <= latestMonth; endMonth++) {
            List<CompanyMonthData> cumulativeList = new ArrayList<>();

            for (int compIdx = 0; compIdx < companyList.size(); compIdx++) {
                String companyCode = companyList.get(compIdx);
                CompanyMonthData cumData = null;

                // 從月份 1 累加到 endMonth
                for (int m = 1; m <= endMonth; m++) {
                    List<CompanyMonthData> mData = monthlyData.get(m);
                    if (mData == null || compIdx >= mData.size()) continue;

                    CompanyMonthData monthCompany = mData.get(compIdx);
                    if (cumData == null) {
                        cumData = new CompanyMonthData(
                                monthCompany.getCompanyCode(),
                                monthCompany.getCompanyName(),
                                monthCompany.getYear(),
                                endMonth);
                        for (String code : CategoryMapping.ALL_INSURANCE_CODES) {
                            cumData.putPremium(code, monthCompany.getPremium(code));
                        }
                    } else {
                        for (String code : CategoryMapping.ALL_INSURANCE_CODES) {
                            cumData.putPremium(code,
                                    cumData.getPremium(code) + monthCompany.getPremium(code));
                        }
                    }
                }

                if (cumData != null) {
                    cumData.recalculateTotal();
                    cumulativeList.add(cumData);
                }
            }

            result.put(endMonth, cumulativeList);
        }

        return result;
    }

    /**
     * 計算累計小計
     */
    public CompanyMonthData calculateSubtotal(List<CompanyMonthData> periodData, int year, int endMonth) {
        CompanyMonthData subtotal = new CompanyMonthData("", "小計", year, endMonth);
        for (String code : CategoryMapping.ALL_INSURANCE_CODES) {
            long sum = periodData.stream().mapToLong(d -> d.getPremium(code)).sum();
            subtotal.putPremium(code, sum);
        }
        subtotal.recalculateTotal();
        return subtotal;
    }

    /**
     * 產生累計期間標籤 (e.g., "1", "1-2", "1-3")
     */
    public static String getPeriodLabel(int endMonth) {
        if (endMonth == 1) return "1";
        return "1-" + endMonth;
    }
}
