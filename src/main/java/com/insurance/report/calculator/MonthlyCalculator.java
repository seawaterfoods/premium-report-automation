package com.insurance.report.calculator;

import com.insurance.report.model.CategoryMapping;
import com.insurance.report.model.CompanyMonthData;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * 單月彙整計算
 * <p>
 * 固定產出 1~12 月份，無資料的月份/公司統一補 0。
 */
@Component
public class MonthlyCalculator {

    /**
     * 將原始資料整理為按「月份 → 公司」結構的單月資料
     * <p>
     * 返回：month(1~12) → List&lt;CompanyMonthData&gt; (依公司代號排序)
     * 無資料的公司/月份以 0 填充。
     *
     * @param rawData     所有來源資料 (今年)
     * @param companyList 公司清單 (已排序)
     */
    public Map<Integer, List<CompanyMonthData>> calculate(
            List<CompanyMonthData> rawData,
            List<String> companyList) {

        // 先將原始資料索引化: (公司代號, 月份) → data
        Map<String, CompanyMonthData> index = new HashMap<>();
        Map<String, String> codeToName = new LinkedHashMap<>();
        for (CompanyMonthData d : rawData) {
            index.put(key(d.getCompanyCode(), d.getMonth()), d);
            codeToName.putIfAbsent(d.getCompanyCode(), d.getCompanyName());
        }

        Map<Integer, List<CompanyMonthData>> result = new LinkedHashMap<>();

        for (int month = 1; month <= 12; month++) {
            List<CompanyMonthData> monthList = new ArrayList<>();
            for (String code : companyList) {
                CompanyMonthData data = index.get(key(code, month));
                if (data != null) {
                    monthList.add(data);
                } else {
                    // 無資料，補 0
                    CompanyMonthData empty = new CompanyMonthData(
                            code, codeToName.getOrDefault(code, code),
                            rawData.isEmpty() ? 0 : rawData.get(0).getYear(), month);
                    for (String insuranceCode : CategoryMapping.ALL_INSURANCE_CODES) {
                        empty.putPremium(insuranceCode, 0);
                    }
                    empty.setTotal(0);
                    monthList.add(empty);
                }
            }
            result.put(month, monthList);
        }

        return result;
    }

    /**
     * 計算單月小計 (所有公司加總)
     */
    public CompanyMonthData calculateSubtotal(List<CompanyMonthData> monthData, int year, int month) {
        CompanyMonthData subtotal = new CompanyMonthData("", "小計", year, month);
        for (String code : CategoryMapping.ALL_INSURANCE_CODES) {
            long sum = monthData.stream().mapToLong(d -> d.getPremium(code)).sum();
            subtotal.putPremium(code, sum);
        }
        subtotal.recalculateTotal();
        return subtotal;
    }

    private String key(String companyCode, int month) {
        return companyCode + "_" + month;
    }
}
