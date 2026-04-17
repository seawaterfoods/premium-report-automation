package com.insurance.report.util;

/**
 * 險種代號正規化工具
 * <p>
 * 來源檔案中 0100~2900 為文字，3000~9900 為數字型態，
 * 需統一轉為 4 位文字 (左補零)。
 */
public final class InsuranceCodeUtil {

    private InsuranceCodeUtil() {
    }

    /**
     * 將險種代號正規化為 4 位文字
     *
     * @param raw 原始值 (可能是 String "0100" 或 Number 3000)
     * @return 正規化後的代號 (e.g., "0100", "3000", "9900")
     */
    public static String normalize(Object raw) {
        if (raw == null) {
            return null;
        }
        String str;
        if (raw instanceof Number) {
            str = String.valueOf(((Number) raw).intValue());
        } else {
            str = raw.toString().trim();
        }
        // 左補零至 4 位
        while (str.length() < 4) {
            str = "0" + str;
        }
        return str;
    }

    /**
     * 檢查是否為合法的險種代號格式 (4 位數字)
     */
    public static boolean isValid(String code) {
        return code != null && code.matches("\\d{4}");
    }
}
