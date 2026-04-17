package com.insurance.report.util;

import com.insurance.report.model.GrowthRateResult;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.math.BigDecimal;
import java.math.RoundingMode;

/**
 * 成長率計算工具
 * <p>
 * 公式：(今年 / 去年) - 1
 * <ul>
 *   <li>去年 = 0 → 以 1 為分母，記錄警告 (缺資料)</li>
 *   <li>去年 < 0 → 以 1 為分母，記錄警告 (負數資料)</li>
 *   <li>結果四捨五入至小數第 2 位</li>
 * </ul>
 */
public final class GrowthRateUtil {

    private static final Logger log = LoggerFactory.getLogger(GrowthRateUtil.class);

    private GrowthRateUtil() {
    }

    /**
     * 計算成長率
     *
     * @param currentValue 今年數值
     * @param priorValue   去年數值
     * @param context      日誌用上下文描述 (e.g., "火險/旺旺友聯/3月")
     * @return 成長率結果 (含警告資訊)
     */
    public static GrowthRateResult calculate(long currentValue, long priorValue, String context) {
        long denominator = priorValue;
        String warning = null;

        if (priorValue == 0) {
            denominator = 1;
            warning = String.format("去年值為 0（缺資料），以 1 為分母計算 [%s]", context);
            log.warn(warning);
        } else if (priorValue < 0) {
            denominator = 1;
            warning = String.format("去年值為負數 (%d)，以 1 為分母計算 [%s]", priorValue, context);
            log.warn(warning);
        }

        double rate = ((double) currentValue / denominator) - 1.0;

        // 四捨五入至小數第 2 位
        BigDecimal bd = BigDecimal.valueOf(rate).setScale(2, RoundingMode.HALF_UP);
        rate = bd.doubleValue();

        if (warning != null) {
            return GrowthRateResult.withWarning(rate, warning);
        }
        return GrowthRateResult.normal(rate);
    }

    /**
     * 格式化成長率為百分比字串 (e.g., "12.34%")
     */
    public static String formatPercent(double rate) {
        BigDecimal pct = BigDecimal.valueOf(rate * 100).setScale(2, RoundingMode.HALF_UP);
        return pct.toPlainString() + "%";
    }
}
