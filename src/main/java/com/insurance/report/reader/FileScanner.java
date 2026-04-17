package com.insurance.report.reader;

import com.insurance.report.config.AppConfig;
import com.insurance.report.model.SourceFileInfo;
import com.insurance.report.model.SourceFileSet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.file.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

/**
 * 掃描 import/ 資料夾，探索所有來源 Excel 檔案
 * <p>
 * 檔名格式：{公司代號}_{民國年月4碼}_保險收入統計表.xlsx
 * 資料夾結構：import/{年份}/{月份}/
 */
@Component
public class FileScanner {

    private static final Logger log = LoggerFactory.getLogger(FileScanner.class);

    /** 檔名正則：公司代號_年月_保險收入統計表.xlsx */
    private static final Pattern FILE_PATTERN =
            Pattern.compile("^(\\d{1,2})_(\\d{3})(\\d{2})_保險收入統計表\\.xlsx$");

    private final AppConfig config;

    public FileScanner(AppConfig config) {
        this.config = config;
    }

    /**
     * 掃描來源資料夾，回傳今年/去年檔案集合
     */
    public SourceFileSet scan() throws IOException {
        int processYear = config.getProcessYear();
        Path importDir = Paths.get(config.getImportDir());

        if (!Files.exists(importDir)) {
            throw new IOException("來源資料夾不存在: " + importDir.toAbsolutePath());
        }

        SourceFileSet fileSet = new SourceFileSet();
        fileSet.setCurrentYear(processYear);
        int priorYear = processYear - 1;

        // 掃描年份資料夾
        Path yearDir = importDir.resolve(String.valueOf(processYear));
        if (!Files.exists(yearDir)) {
            throw new IOException("年份資料夾不存在: " + yearDir.toAbsolutePath());
        }

        log.info("掃描來源資料夾: {}", yearDir.toAbsolutePath());

        // 遍歷月份子資料夾
        try (Stream<Path> monthDirs = Files.list(yearDir)) {
            monthDirs.filter(Files::isDirectory)
                    .sorted()
                    .forEach(monthDir -> scanMonthDir(monthDir, processYear, priorYear, fileSet));
        }

        log.info("掃描完成: {}", fileSet);
        return fileSet;
    }

    private void scanMonthDir(Path monthDir, int currentYear, int priorYear, SourceFileSet fileSet) {
        try (Stream<Path> files = Files.list(monthDir)) {
            files.filter(p -> p.toString().endsWith(".xlsx"))
                    .forEach(file -> {
                        String fileName = file.getFileName().toString();
                        Matcher matcher = FILE_PATTERN.matcher(fileName);
                        if (!matcher.matches()) {
                            log.warn("檔名格式不符，跳過: {}", fileName);
                            return;
                        }

                        String companyCode = matcher.group(1);
                        int fileYear = Integer.parseInt(matcher.group(2));
                        int fileMonth = Integer.parseInt(matcher.group(3));

                        // 公司代號左補零至 2 位
                        if (companyCode.length() == 1) {
                            companyCode = "0" + companyCode;
                        }

                        SourceFileInfo info = new SourceFileInfo(file, companyCode, fileYear, fileMonth);

                        if (fileYear == currentYear) {
                            fileSet.addCurrentYearFile(info);
                            log.debug("今年檔案: {}", info);
                        } else if (fileYear == priorYear) {
                            fileSet.addPriorYearFile(info);
                            log.debug("去年檔案: {}", info);
                        } else {
                            log.warn("年份不在處理範圍 ({}/{}), 跳過: {}", currentYear, priorYear, fileName);
                        }
                    });
        } catch (IOException e) {
            log.error("掃描月份資料夾失敗: {}", monthDir, e);
        }
    }
}
