package com.insurance.report.model;

import java.util.*;

/**
 * 歸屬對照表 — 險種代號 → 九大類的分類規則
 * <p>
 * 支援自動擴展：若來源檔案中出現新險種，可動態新增。
 */
public class CategoryMapping {

    /** 九大類名稱 (順序固定) */
    public static final List<String> MAJOR_CATEGORIES = List.of(
            "火險", "水險", "航空險", "汽車險", "意外險",
            "傷害險", "天災險", "健康險", "國外分進"
    );

    /**
     * 15 子分類定義 (用於「總」表和比較表的欄位)
     * 每個子類包含：顯示名稱、歸屬險種代號清單
     */
    public static final List<SubCategory> SUB_CATEGORIES;

    static {
        List<SubCategory> list = new ArrayList<>();
        list.add(new SubCategory("火險", "火險", List.of("0100", "0200", "0300", "0400")));
        list.add(new SubCategory("水險", "水險", List.of("0500", "0600", "0700", "0800")));
        list.add(new SubCategory("航空險", "航空險", List.of("0900")));
        list.add(new SubCategory("車體損失險", "汽車險", List.of("1000", "1100")));
        list.add(new SubCategory("任意責任險", "汽車險", List.of("1200", "1300")));
        list.add(new SubCategory("強制責任-汽車", "汽車險", List.of("1400", "1500")));
        list.add(new SubCategory("強制-機車", "汽車險", List.of("1600")));
        list.add(new SubCategory("強制-電動二輪", "汽車險", List.of("3200")));
        list.add(new SubCategory("責任險", "意外險", List.of("1700", "1800")));
        list.add(new SubCategory("工程險", "意外險", List.of("1900")));
        list.add(new SubCategory("信用保證", "意外險", List.of("2100", "2200")));
        list.add(new SubCategory("其他財產責任保險", "意外險", List.of("2000", "2300", "2600", "2700")));
        list.add(new SubCategory("傷害險", "傷害險", List.of("2400")));
        list.add(new SubCategory("天災險", "天災險", List.of("2500", "2800", "2900")));
        list.add(new SubCategory("健康險", "健康險", List.of("3000", "3100")));
        SUB_CATEGORIES = Collections.unmodifiableList(list);
    }

    /** 國外分進 (獨立處理，可由 config 控制) */
    public static final SubCategory OVERSEAS_REINSURANCE =
            new SubCategory("國外分進", "國外分進", List.of("9900"));

    /**
     * 33 種險種代號的完整順序 (對應「單」表 D~AJ 欄)
     */
    public static final List<String> ALL_INSURANCE_CODES = List.of(
            "0100", "0200", "0300", "0400",
            "0500", "0600", "0700", "0800",
            "0900",
            "1000", "1100", "1200", "1300",
            "1400", "1500", "1600",
            "1700", "1800", "1900",
            "2000", "2100", "2200", "2300",
            "2400", "2500", "2600", "2700",
            "2800", "2900",
            "3000", "3100", "3200",
            "9900"
    );

    /**
     * 險種代號 → 險種短名 (用於表頭)
     */
    public static final Map<String, String> CODE_TO_SHORT_NAME;

    static {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("0100", "一年期\n住宅火險");
        map.put("0200", "長期\n住宅火險");
        map.put("0300", "一年期\n商業火險");
        map.put("0400", "長期\n商業火險");
        map.put("0500", "內陸運輸險");
        map.put("0600", "貨物運輸險");
        map.put("0700", "船體險");
        map.put("0800", "漁船險");
        map.put("0900", "航空險");
        map.put("1000", "自用車損險");
        map.put("1100", "商業車損險");
        map.put("1200", "自用車責險");
        map.put("1300", "商業車責險");
        map.put("1400", "強制自用車責險");
        map.put("1500", "強制商業車責險");
        map.put("1600", "強制機車責險");
        map.put("1700", "一般責任險");
        map.put("1800", "專業責任險");
        map.put("1900", "工程險");
        map.put("2000", "核能險");
        map.put("2100", "保證險");
        map.put("2200", "信用險");
        map.put("2300", "其他財產保險");
        map.put("2400", "傷害險");
        map.put("2500", "商業地震險");
        map.put("2600", "個人綜險");
        map.put("2700", "商業綜險");
        map.put("2800", "颱風洪水險");
        map.put("2900", "政策地震險");
        map.put("3000", "一年期\n健康險");
        map.put("3100", "長年期\n健康險");
        map.put("3200", "強制微型電動\n二輪車責險");
        map.put("9900", "國外\n分進");
        CODE_TO_SHORT_NAME = Collections.unmodifiableMap(map);
    }

    /**
     * 險種代號 → 險種全名 (用於歸屬表)
     */
    public static final Map<String, String> CODE_TO_FULL_NAME;

    static {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("0100", "一年期住宅火災保險");
        map.put("0200", "長期住宅火災保險");
        map.put("0300", "一年期商業火災保險");
        map.put("0400", "長期商業火災保險");
        map.put("0500", "內陸運輸保險");
        map.put("0600", "貨物運輸保險");
        map.put("0700", "船體保險");
        map.put("0800", "漁船保險");
        map.put("0900", "航空保險");
        map.put("1000", "一般自用汽車財產損失險");
        map.put("1100", "一般商業汽車財產損失險");
        map.put("1200", "一般自用汽車責任險");
        map.put("1300", "一般商業汽車責任險");
        map.put("1400", "強制自用汽車責任險");
        map.put("1500", "強制商業汽車責任險");
        map.put("1600", "強制機車責任險");
        map.put("1700", "一般責任保險");
        map.put("1800", "專業責任保險");
        map.put("1900", "工程保險");
        map.put("2000", "核能保險");
        map.put("2100", "保證保險");
        map.put("2200", "信用保險");
        map.put("2300", "其他財產保險");
        map.put("2400", "傷害險");
        map.put("2500", "商業性地震險");
        map.put("2600", "個人綜合保險");
        map.put("2700", "商業綜合保險");
        map.put("2800", "颱風、洪水保險");
        map.put("2900", "政策性地震險");
        map.put("3000", "一年期健康險");
        map.put("3100", "長年期健康險");
        map.put("3200", "強制微型電動二輪車責任險");
        map.put("9900", "國外分進業務");
        CODE_TO_FULL_NAME = Collections.unmodifiableMap(map);
    }

    /** 九大類序號 (用於歸屬表「類」欄) */
    public static final Map<String, String> MAJOR_CATEGORY_NUMBER = Map.of(
            "火險", "一", "水險", "二", "航空險", "三",
            "汽車險", "四", "意外險", "五", "傷害險", "六",
            "天災險", "七", "健康險", "八", "國外分進", "九"
    );

    /**
     * 險種代號 → 子分組顯示名 (用於歸屬表 D 欄)
     * 僅在子分組第一個代號處填入名稱
     */
    public static final Map<String, String> CODE_TO_SUB_GROUP;

    static {
        Map<String, String> map = new LinkedHashMap<>();
        map.put("1000", "車體損");
        map.put("1200", "任意車責");
        map.put("1400", "強制車");
        map.put("1600", "強制機");
        map.put("1700", "意責任");
        map.put("1900", "意工");
        map.put("2100", "意信保");
        map.put("2000", "意其他");
        CODE_TO_SUB_GROUP = Collections.unmodifiableMap(map);
    }

    /**
     * 取得全部 16 子分類 (含國外分進)
     */
    public static List<SubCategory> getAllSubCategoriesWithOverseas() {
        List<SubCategory> list = new ArrayList<>(SUB_CATEGORIES);
        list.add(OVERSEAS_REINSURANCE);
        return list;
    }

    /**
     * 子分類定義
     */
    public static class SubCategory {
        private final String name;
        private final String majorCategory;
        private final List<String> insuranceCodes;

        public SubCategory(String name, String majorCategory, List<String> insuranceCodes) {
            this.name = name;
            this.majorCategory = majorCategory;
            this.insuranceCodes = Collections.unmodifiableList(new ArrayList<>(insuranceCodes));
        }

        public String getName() {
            return name;
        }

        public String getMajorCategory() {
            return majorCategory;
        }

        public List<String> getInsuranceCodes() {
            return insuranceCodes;
        }

        /**
         * 從公司月資料中加總此子分類的所有險種金額
         */
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
