package com.insurance.report.calculator;

import com.insurance.report.model.CompanyInfo;
import org.springframework.stereotype.Component;

import java.util.*;

/**
 * 公司排名計算
 * <p>
 * 取年度累計至最新月份的保費合計，由大到小排序。
 */
@Component
public class RankingCalculator {

    /**
     * 排名結果
     */
    public static class RankEntry {
        private final int rank;
        private final CompanyInfo company;
        private final long totalPremium;

        public RankEntry(int rank, CompanyInfo company, long totalPremium) {
            this.rank = rank;
            this.company = company;
            this.totalPremium = totalPremium;
        }

        public int getRank() { return rank; }
        public CompanyInfo getCompany() { return company; }
        public long getTotalPremium() { return totalPremium; }
    }

    /**
     * 計算排名
     *
     * @param companyTotals 公司 → 累計保費合計
     * @return 排名清單 (由大到小)
     */
    public List<RankEntry> calculate(Map<CompanyInfo, Long> companyTotals) {
        List<Map.Entry<CompanyInfo, Long>> sorted = new ArrayList<>(companyTotals.entrySet());
        sorted.sort((a, b) -> Long.compare(b.getValue(), a.getValue()));

        List<RankEntry> result = new ArrayList<>();
        for (int i = 0; i < sorted.size(); i++) {
            result.add(new RankEntry(i + 1, sorted.get(i).getKey(), sorted.get(i).getValue()));
        }
        return result;
    }
}
