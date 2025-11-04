package com.tsb.noise.service.utils;

import com.tsb.noise.model.FileType;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class FileUtils {
    private static final Logger log = LoggerFactory.getLogger(FileUtils.class);

    private FileUtils() {
        // Utility class
    }

    /**
     * Находит все поддерживаемые Excel файлы УЗД
     */
    public static List<File> findTargetExcelFiles(String rootPath) throws IOException {
        if (rootPath == null || rootPath.trim().isEmpty()) {
            log.warn("Путь для поиска файлов не указан");
            return Collections.emptyList();
        }

        Path start = Paths.get(rootPath);

        if (!Files.exists(start)) {
            log.error("Путь не существует: {}", rootPath);
            throw new IllegalArgumentException("Путь не существует: " + rootPath);
        }

        if (!Files.isDirectory(start)) {
            log.error("Указанный путь не является директорией: {}", rootPath);
            throw new IllegalArgumentException("Указанный путь не является директорией: " + rootPath);
        }

        try (Stream<Path> stream = Files.walk(start)) {
            List<File> foundFiles = stream
                    .filter(FileUtils::isSupportedExcelFile)
                    .map(Path::toFile)
                    .collect(Collectors.toList());

            log.info("Найдено поддерживаемых файлов УЗД: {}", foundFiles.size());

            // Логируем типы найденных файлов
            foundFiles.forEach(file -> {
                FileType fileType = FileType.fromFileName(file.getName());
                if (fileType != null) {
                    log.debug("Найден файл типа {}: {}", fileType.getDisplayName(), file.getName());
                }
            });

            return foundFiles;
        } catch (IOException e) {
            log.error("Ошибка при поиске файлов в директории {}: {}", rootPath, e.getMessage(), e);
            throw e;
        }
    }

    /**
     * Проверяет является ли файл поддерживаемым Excel файлом УЗД
     */
    private static boolean isSupportedExcelFile(Path path) {
        if (!Files.isRegularFile(path)) {
            return false;
        }

        String fileName = path.getFileName().toString();

        // Проверяем что файл Excel и соответствует одному из шаблонов
        boolean isExcel = fileName.endsWith(".xlsx") || fileName.endsWith(".xls");
        boolean isSupported = FileType.isSupportedFile(fileName);

        if (isSupported && isExcel) {
            log.debug("Найден поддерживаемый файл: {}", fileName);
        }

        return isSupported && isExcel;
    }

    /**
     * Генерирует имя выходного файла на основе типа исходного файла
     */
    public static String generateOutputFileName(String inputFileName) {
        FileType fileType = FileType.fromFileName(inputFileName);
        if (fileType != null) {
            String outputName = inputFileName.replace(fileType.getInputPattern(), fileType.getOutputPattern());
            log.debug("Сгенерировано имя выходного файла: {} -> {}", inputFileName, outputName);
            return outputName;
        }

        // Fallback для обратной совместимости
        String outputName = inputFileName.replace("УЗД в РТ ТХ день_FINAL", "В записку_УЗД в РТ ТХ день_FINAL");
        log.debug("Сгенерировано имя выходного файла (fallback): {} -> {}", inputFileName, outputName);
        return outputName;
    }

    /**
     * Получает отображаемое имя типа файла
     */
    public static String getFileTypeDisplayName(String fileName) {
        FileType fileType = FileType.fromFileName(fileName);
        return fileType != null ? fileType.getDisplayName() : "Неизвестный тип";
    }

    /**
     * Проверяет, содержит ли файл необходимый лист
     */
    public static boolean hasRequiredSheet(File file) {
        try {
            // Быстрая проверка без полной загрузки файла
            if (!file.exists() || !file.canRead()) {
                return false;
            }

            // Более детальная проверка будет в ExcelProcessor
            return true;
        } catch (Exception e) {
            log.warn("Ошибка при проверке файла {}: {}", file.getName(), e.getMessage());
            return false;
        }
    }
}