package com.tsb.noise.service.utils;

import com.tsb.noise.model.FileType;
import com.tsb.noise.model.ProcessConfig;
import com.tsb.noise.service.processors.BaseExcelProcessor;
import com.tsb.noise.service.processors.ProcessorFactory;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.util.List;
import java.util.function.Consumer;

@Slf4j
public class FileProcessorCoordinator {

    public Void processFiles(String rootPath, ProcessConfig config,
                             Consumer<Double> progressCallback, Consumer<String> messageCallback) {
        try {
            messageCallback.accept("🔍 Поиск файлов...");
            progressCallback.accept(0.0);

            List<File> targetFiles = FileUtils.findTargetExcelFiles(rootPath);

            if (targetFiles.isEmpty()) {
                messageCallback.accept("❌ Файлы не найдены");
                progressCallback.accept(100.0);
                return null;
            }

            messageCallback.accept("✅ Найдено файлов: " + targetFiles.size());
            progressCallback.accept(10.0);

            int processed = 0;
            int totalFiles = targetFiles.size();
            int successful = 0;
            int failed = 0;

            for (File inputFile : targetFiles) {
                String currentFileName = inputFile.getName();
                messageCallback.accept("Обработка " + (processed + 1) + "/" + totalFiles + ": " + currentFileName);

                try {
                    FileType fileType = FileType.fromFileName(currentFileName);
                    BaseExcelProcessor processor = ProcessorFactory.createProcessor(fileType);

                    ProcessConfig fileConfig = new ProcessConfig(
                            config.isRemoveSoundIsolation(),
                            config.isMoveSoundIsolation(),
                            fileType
                    );

                    String outputFileName = FileUtils.generateOutputFileName(currentFileName);
                    File outputFile = new File(inputFile.getParent(), outputFileName);

                    boolean success = processor.process(inputFile, outputFile, fileConfig);

                    if (success) {
                        successful++;
                        log.info("✅ Успешно обработан: {}", outputFile.getName());
                    } else {
                        failed++;
                        log.error("❌ Ошибка при обработке: {}", inputFile.getName());
                    }

                } catch (Exception e) {
                    failed++;
                    log.error("❌ Критическая ошибка при обработке {}: {}", currentFileName, e.getMessage(), e);
                }

                processed++;
                double progress = 10.0 + (processed * 90.0 / totalFiles);
                progressCallback.accept(progress);

                try {
                    Thread.sleep(50);
                } catch (InterruptedException ie) {
                    Thread.currentThread().interrupt();
                    break;
                }
            }

            String resultMessage = String.format("🎉 Обработка завершена! Успешно: %d, Ошибок: %d", successful, failed);
            if (config.isRemoveSoundIsolation() || config.isMoveSoundIsolation()) {
                resultMessage += getOperationsSummary(config);
            }

            messageCallback.accept(resultMessage);
            progressCallback.accept(100.0);
            log.info(resultMessage);

        } catch (Exception e) {
            log.error("💥 Критическая ошибка при обработке файлов: {}", e.getMessage(), e);
            messageCallback.accept("💥 Критическая ошибка: " + e.getMessage());
            throw new RuntimeException("Ошибка обработки файлов", e);
        }

        return null;
    }

    private String getOperationsSummary(ProcessConfig config) {
        StringBuilder summary = new StringBuilder(" (");

        if (config.isRemoveSoundIsolation()) {
            summary.append("удалена звукоизоляция");
        }

        if (config.isRemoveSoundIsolation() && config.isMoveSoundIsolation()) {
            summary.append(", ");
        }

        if (config.isMoveSoundIsolation()) {
            summary.append("перемещены преграды");
        }

        summary.append(")");
        return summary.toString();
    }

    public boolean canProcessFiles(String rootPath) {
        try {
            List<File> targetFiles = FileUtils.findTargetExcelFiles(rootPath);
            return !targetFiles.isEmpty();
        } catch (Exception e) {
            log.error("Ошибка проверки файлов: {}", e.getMessage(), e);
            return false;
        }
    }

    public int getFileCount(String rootPath) {
        try {
            List<File> targetFiles = FileUtils.findTargetExcelFiles(rootPath);
            return targetFiles.size();
        } catch (Exception e) {
            log.error("Ошибка подсчета файлов: {}", e.getMessage(), e);
            return 0;
        }
    }
}