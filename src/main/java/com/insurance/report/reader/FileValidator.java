package com.insurance.report.reader;

import com.insurance.report.model.SourceFileInfo;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

/**
 * 驗證來源檔案結構
 */
@Component
public class FileValidator {

    private static final Logger log = LoggerFactory.getLogger(FileValidator.class);

    /**
     * 驗證檔案是否存在且可讀
     *
     * @return 驗證錯誤清單 (空表示通過)
     */
    public List<String> validate(SourceFileInfo fileInfo) {
        List<String> errors = new ArrayList<>();

        if (!Files.exists(fileInfo.getFilePath())) {
            errors.add("檔案不存在: " + fileInfo.getFilePath());
            return errors;
        }

        if (!Files.isReadable(fileInfo.getFilePath())) {
            errors.add("檔案無法讀取: " + fileInfo.getFilePath());
            return errors;
        }

        try {
            long size = Files.size(fileInfo.getFilePath());
            if (size == 0) {
                errors.add("檔案大小為 0: " + fileInfo.getFilePath());
            }
        } catch (IOException e) {
            errors.add("無法讀取檔案大小: " + fileInfo.getFilePath());
        }

        return errors;
    }
}
