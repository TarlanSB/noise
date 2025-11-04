package com.tsb.noise.service;

import lombok.Data;

import java.util.List;

/**
 * Результат обработки файлов
 */
@Data
public class ProcessingResult {
    private boolean success;
    private int totalFiles;
    private int processedFiles;
    private int failedFiles;
    private List<String> processedFileNames;
    private List<String> errorMessages;
    private String summary;

    public ProcessingResult() {
        this.success = true;
    }
}