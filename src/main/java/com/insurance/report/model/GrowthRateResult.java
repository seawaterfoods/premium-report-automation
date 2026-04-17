package com.insurance.report.model;

/**
 * 成長率計算結果
 */
public class GrowthRateResult {

    private final double rate;
    private final boolean warning;
    private final String warningMessage;

    private GrowthRateResult(double rate, boolean warning, String warningMessage) {
        this.rate = rate;
        this.warning = warning;
        this.warningMessage = warningMessage;
    }

    public static GrowthRateResult normal(double rate) {
        return new GrowthRateResult(rate, false, null);
    }

    public static GrowthRateResult withWarning(double rate, String warningMessage) {
        return new GrowthRateResult(rate, true, warningMessage);
    }

    public double getRate() {
        return rate;
    }

    public boolean hasWarning() {
        return warning;
    }

    public String getWarningMessage() {
        return warningMessage;
    }

    @Override
    public String toString() {
        String s = String.format("%.2f%%", rate * 100);
        if (warning) {
            s += " [⚠ " + warningMessage + "]";
        }
        return s;
    }
}
