package com.insurance.report.model;

import java.util.*;

/**
 * 掃描結果 — 今年和去年的來源檔案集合
 */
public class SourceFileSet {

    /** 今年來源檔（按月份分組） */
    private final Map<Integer, List<SourceFileInfo>> currentYearFiles = new TreeMap<>();
    /** 去年來源檔（按月份分組） */
    private final Map<Integer, List<SourceFileInfo>> priorYearFiles = new TreeMap<>();

    private int currentYear;
    private int priorYear;
    private int latestMonth;

    public void addCurrentYearFile(SourceFileInfo file) {
        currentYearFiles.computeIfAbsent(file.getMonth(), k -> new ArrayList<>()).add(file);
        latestMonth = Math.max(latestMonth, file.getMonth());
    }

    public void addPriorYearFile(SourceFileInfo file) {
        priorYearFiles.computeIfAbsent(file.getMonth(), k -> new ArrayList<>()).add(file);
    }

    public Map<Integer, List<SourceFileInfo>> getCurrentYearFiles() {
        return currentYearFiles;
    }

    public Map<Integer, List<SourceFileInfo>> getPriorYearFiles() {
        return priorYearFiles;
    }

    public int getCurrentYear() {
        return currentYear;
    }

    public void setCurrentYear(int currentYear) {
        this.currentYear = currentYear;
        this.priorYear = currentYear - 1;
    }

    public int getPriorYear() {
        return priorYear;
    }

    /** 有資料的最新月份 */
    public int getLatestMonth() {
        return latestMonth;
    }

    /** 今年有幾家公司的資料 */
    public Set<String> getAllCompanyCodes() {
        Set<String> codes = new TreeSet<>();
        currentYearFiles.values().stream()
                .flatMap(Collection::stream)
                .forEach(f -> codes.add(f.getCompanyCode()));
        return codes;
    }

    /** 今年有資料的月份清單 */
    public Set<Integer> getMonthsWithData() {
        return currentYearFiles.keySet();
    }

    public int getTotalFileCount() {
        int count = currentYearFiles.values().stream().mapToInt(List::size).sum();
        count += priorYearFiles.values().stream().mapToInt(List::size).sum();
        return count;
    }

    @Override
    public String toString() {
        return String.format("SourceFileSet{year=%d, months=%s, companies=%d, totalFiles=%d}",
                currentYear, getMonthsWithData(), getAllCompanyCodes().size(), getTotalFileCount());
    }
}
