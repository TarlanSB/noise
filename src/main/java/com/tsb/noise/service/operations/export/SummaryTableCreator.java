package com.tsb.noise.service.operations.export;

import com.tsb.noise.model.FileType;
import com.tsb.noise.service.operations.core.SheetLayoutManager;
import com.tsb.noise.service.operations.core.StyleApplier;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Создатель сводной таблицы расчетных точек с новой структурой (блоки по 3 строки)
 */
@Slf4j
public class SummaryTableCreator implements TableCreator {

    private final StyleApplier styleApplier;
    private final SheetLayoutManager layoutManager;

    public SummaryTableCreator() {
        this.styleApplier = new StyleApplier();
        this.layoutManager = new SheetLayoutManager();
    }

    /**
     * Создает сводную таблицу РТ с новой структурой
     */
    @Override
    public boolean createTable(String rootPath, boolean createSummaryTable) {
        if (!createSummaryTable) {
            log.info("Создание сводной таблицы отключено");
            return false;
        }

        log.info("🚀 Начало создания сводной таблицы РТ с новой структурой...");

        try {
            // Находим все файлы для обработки
            List<File> sourceFiles = findAllSourceFiles(rootPath);
            if (sourceFiles.isEmpty()) {
                log.error("❌ Не найдены файлы для создания сводной таблицы");
                return false;
            }

            log.info("✅ Найдено файлов для обработки: {}", sourceFiles.size());

            // Сортируем файлы по номеру ШК
            List<File> sortedFiles = sortFilesByShk(sourceFiles);

            // Извлекаем уникальные РТ из всех файлов
            Set<String> uniqueRtNames = extractUniqueRtNames(sortedFiles);
            if (uniqueRtNames.isEmpty()) {
                log.warn("⚠️ Не найдены РТ для сводной таблицы");
                return false;
            }

            log.info("✅ Найдено уникальных РТ: {}", uniqueRtNames.size());

            // Создаем папку и файл
            File outputFolder = createOutputFolder(rootPath);
            File outputFile = createOutputFile(outputFolder);

            // Создаем сводную таблицу с новой структурой
            return createNewStructureWorkbook(sortedFiles, new ArrayList<>(uniqueRtNames), outputFile);

        } catch (Exception e) {
            log.error("❌ Ошибка при создании сводной таблицы: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * Сортирует файлы по номеру ШК
     */
    private List<File> sortFilesByShk(List<File> files) {
        List<File> sortedFiles = new ArrayList<>(files);
        sortedFiles.sort((f1, f2) -> {
            int shk1 = extractShkNumericValue(f1.getName());
            int shk2 = extractShkNumericValue(f2.getName());
            return Integer.compare(shk1, shk2);
        });

        log.debug("Файлы отсортированы по ШК: {}",
                sortedFiles.stream().map(f -> extractShkNumber(f.getName())).toList());
        return sortedFiles;
    }

    /**
     * Извлекает номер ШК из имени файла
     */
    private String extractShkNumber(String fileName) {
        Pattern pattern = Pattern.compile("ШК(\\d+п?)");
        Matcher matcher = pattern.matcher(fileName);
        return matcher.find() ? "ШК" + matcher.group(1) : "ШК1";
    }

    /**
     * Извлекает числовое значение ШК для сортировки
     */
    private int extractShkNumericValue(String fileName) {
        Pattern pattern = Pattern.compile("ШК(\\d+)");
        Matcher matcher = pattern.matcher(fileName);
        return matcher.find() ? Integer.parseInt(matcher.group(1)) : 1;
    }

    /**
     * Извлекает уникальные наименования РТ из всех файлов
     */
    private Set<String> extractUniqueRtNames(List<File> files) {
        Set<String> uniqueRtNames = new TreeSet<>(this::compareRtNames);

        for (File file : files) {
            try (FileInputStream fis = new FileInputStream(file);
                 Workbook workbook = WorkbookFactory.create(fis)) {

                Sheet sheet = workbook.getSheet("ЛИСТ2");
                if (sheet == null) continue;

                extractRtNamesFromSheet(sheet, uniqueRtNames);

            } catch (Exception e) {
                log.warn("⚠️ Не удалось извлечь РТ из файла {}: {}", file.getName(), e.getMessage());
            }
        }

        return uniqueRtNames;
    }

    /**
     * Компаратор для сортировки РТ
     */
    private int compareRtNames(String rt1, String rt2) {
        // Сначала числовые РТ (РТ-1, РТ-2...), потом РТ-13К, РТ-14К...
        boolean isRt1Numeric = rt1.matches("РТ-\\d+$");
        boolean isRt2Numeric = rt2.matches("РТ-\\d+$");

        if (isRt1Numeric && !isRt2Numeric) return -1;
        if (!isRt1Numeric && isRt2Numeric) return 1;

        return rt1.compareTo(rt2);
    }

    /**
     * Извлекает наименования РТ из листа
     */
    private void extractRtNamesFromSheet(Sheet sheet, Set<String> uniqueRtNames) {
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell cellA = row.getCell(0); // Наименование РТ
            Cell cellB = row.getCell(1); // Тип данных

            if (isRtRow(cellA, cellB)) {
                String rtName = getCellStringValue(cellA).trim();
                if (!rtName.isEmpty()) {
                    uniqueRtNames.add(rtName);
                }
            }
        }
    }

    /**
     * Создает рабочую книгу с новой структурой
     */
    private boolean createNewStructureWorkbook(List<File> sortedFiles, List<String> rtNames, File outputFile) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Сводная таблица УЗД");

            // Настраиваем layout
            setupNewStructureLayout(sheet, rtNames.size());

            // Создаем шапку таблицы с РТ
            createNewStructureHeader(workbook, sheet, rtNames);

            // Заполняем данные из файлов
            fillNewStructureData(workbook, sheet, sortedFiles, rtNames);

            // Сохраняем файл
            outputFile.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            log.info("✅ Успешно создана сводная таблица с новой структурой: {}", outputFile.getAbsolutePath());
            return true;

        } catch (IOException e) {
            log.error("❌ Ошибка при создании сводной таблицы: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * Настраивает layout для новой структуры
     */
    private void setupNewStructureLayout(Sheet sheet, int numRt) {
        // Ширина колонки A (фиксированные заголовки)
        sheet.setColumnWidth(0, 4000);

        // Ширина колонок с данными РТ
        int dataColumnWidth = 2000;
        for (int i = 1; i <= numRt + 1; i++) {
            sheet.setColumnWidth(i, dataColumnWidth);
        }

        sheet.setDefaultRowHeightInPoints(20);
    }

    /**
     * Создает шапку таблицы с новой структурой
     */
    private void createNewStructureHeader(Workbook workbook, Sheet sheet, List<String> rtNames) {
        // Строка 1: "Расчетная точка (РТ)"
        Row row1 = sheet.createRow(0);
        Cell cellA1 = row1.createCell(0);
        cellA1.setCellValue("Расчетная точка (РТ)");

        // Заполняем наименования РТ
        for (int i = 0; i < rtNames.size(); i++) {
            Cell cell = row1.createCell(i + 1);
            cell.setCellValue(rtNames.get(i));
        }

        // Строка 2: "Отметка, м"
        Row row2 = sheet.createRow(1);
        Cell cellA2 = row2.createCell(0);
        cellA2.setCellValue("Отметка, м");

        // Строка 3: "Тип территории"
        Row row3 = sheet.createRow(2);
        Cell cellA3 = row3.createCell(0);
        cellA3.setCellValue("Тип территории");

        // Применяем стили к шапке
        applyNewHeaderStyles(workbook, sheet, 0, 2, rtNames.size() + 1);
    }

    /**
     * Применяет стили к новой шапке
     */
    private void applyNewHeaderStyles(Workbook workbook, Sheet sheet, int startRow, int endRow, int numColumns) {
        CellStyle headerStyle = createHeaderStyle(workbook);

        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                for (int colNum = 0; colNum < numColumns; colNum++) {
                    Cell cell = row.getCell(colNum);
                    if (cell != null) {
                        cell.setCellStyle(headerStyle);
                    }
                }
            }
        }
    }

    /**
     * Заполняет данные с новой структурой
     */
    private void fillNewStructureData(Workbook workbook, Sheet sheet, List<File> files, List<String> rtNames) {
        int currentRow = 3;

        log.info("🔍 НАЧАЛО ЗАПОЛНЕНИЯ ДАННЫХ");
        log.info("📋 Файлов для обработки: {}", files.size());

        // Логируем все файлы
        for (int i = 0; i < files.size(); i++) {
            File file = files.get(i);
            log.info("   {}. {} -> ШК{}", i + 1, file.getName(), extractShkNumber(file.getName()));
        }

        for (File file : files) {
            try {
                log.info("🔄 ОБРАБОТКА: {} в строке {}", file.getName(), currentRow);

                // Проверяем, что строка свободна
                if (currentRow <= sheet.getLastRowNum()) {
                    Row existingRow = sheet.getRow(currentRow);
                    if (existingRow != null) {
                        log.warn("⚠️ Строка {} уже занята! Содержимое: {}", currentRow,
                                getRowDebugInfo(existingRow));
                    }
                }

                currentRow = processFileBlock(workbook, sheet, file, rtNames, currentRow);

            } catch (Exception e) {
                log.warn("⚠️ Ошибка при обработке файла {}: {}", file.getName(), e.getMessage());
            }
        }

        log.info("✅ ЗАВЕРШЕНО. Всего строк: {}", currentRow - 3);
    }

    /**
     * Отладочная информация о строке
     */
    private String getRowDebugInfo(Row row) {
        if (row == null) return "null";
        StringBuilder info = new StringBuilder();
        for (int i = 0; i < Math.min(5, row.getLastCellNum()); i++) {
            Cell cell = row.getCell(i);
            if (cell != null) {
                info.append("[").append(i).append(":").append(getCellStringValue(cell)).append("] ");
            }
        }
        return info.toString();
    }
    /**
     * Обрабатывает блок данных из одного файла (3 строки)
     */
    /**
     * Обрабатывает блок данных из одного файла (3 строки С ДАННЫМИ)
     */
    private int processFileBlock(Workbook workbook, Sheet sheet, File file, List<String> rtNames, int startRow) {
        String fileName = file.getName();
        String shkNumber = extractShkNumber(fileName);
        FileType fileType = FileType.fromFileName(fileName);

        if (fileType == null) {
            return startRow;
        }

        String timeOfDay = fileType.getDisplayName().contains("ночь") ? "ночь" : "день";
        String blockHeader = shkNumber + ", " + timeOfDay;

        log.info("📊 Обработка блока: {} -> {}", fileName, blockHeader);

        // Создаем 3 строки для блока С ДАННЫМИ
        Row noiseRow = sheet.createRow(startRow);
        Row pduRow = sheet.createRow(startRow + 1);
        Row excessRow = sheet.createRow(startRow + 2);

        // Заполняем заголовки в колонке A
        noiseRow.createCell(0).setCellValue(blockHeader);
        pduRow.createCell(0).setCellValue("ПДУ");
        excessRow.createCell(0).setCellValue("Превышение");

        // Извлекаем данные из файла
        try (FileInputStream fis = new FileInputStream(file);
             Workbook fileWorkbook = WorkbookFactory.create(fis)) {

            Sheet fileSheet = fileWorkbook.getSheet("ЛИСТ2");
            if (fileSheet != null) {
                Map<String, FileData> fileData = extractFileData(fileSheet, rtNames);

                // ЗАПОЛНЯЕМ ДАННЫЕ В ТЕ ЖЕ САМЫЕ СТРОКИ
                for (int i = 0; i < rtNames.size(); i++) {
                    String rtName = rtNames.get(i);
                    FileData data = fileData.get(rtName);
                    int colIndex = i + 1;

                    if (data != null) {
                        // УЗД данные в ПЕРВУЮ строку блока
                        if (data.noiseLevel != null) {
                            noiseRow.createCell(colIndex).setCellValue(data.noiseLevel);
                        }

                        // ПДУ значения во ВТОРУЮ строку блока
                        if (data.pduValue != null) {
                            pduRow.createCell(colIndex).setCellValue(data.pduValue);
                        }

                        // Превышения в ТРЕТЬЮ строку блока
                        if (data.excessValue != null) {
                            excessRow.createCell(colIndex).setCellValue(data.excessValue);
                        }
                    }
                }
            }

        } catch (Exception e) {
            log.warn("⚠️ Ошибка при извлечении данных из файла {}: {}", fileName, e.getMessage());
        }

        // Применяем стили к блоку
        applyBlockStyles(workbook, sheet, startRow, startRow + 2, rtNames.size() + 1);

        return startRow + 3; // Переходим к следующему блоку
    }
    /**
     * Извлекает данные из файла с учетом группировки по РТ через пустые строки
     */
    private Map<String, FileData> extractFileData(Sheet sheet, List<String> rtNames) {
        Map<String, FileData> fileData = new HashMap<>();
        String currentRt = null;

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell cellA = row.getCell(0); // Столбец A - наименование РТ
            Cell cellB = row.getCell(1); // Столбец B - тип данных
            Cell cellL = row.getCell(11); // Столбец L - значение

            // Проверяем, является ли строка началом новой группы РТ
            if (isNewRtGroup(cellA, cellB)) {
                currentRt = getCellStringValue(cellA).trim();

                // Если это новая РТ из нашего списка, создаем для нее запись
                if (rtNames.contains(currentRt)) {
                    fileData.putIfAbsent(currentRt, new FileData());
                } else {
                    currentRt = null; // Пропускаем РТ не из списка
                }
            }

            // Если мы внутри группы РТ, обрабатываем данные
            if (currentRt != null && cellB != null) {
                processDataRow(fileData.get(currentRt), cellB, cellL);
            }
        }

        return fileData;
    }

    /**
     * Проверяет, является ли строка началом новой группы РТ
     */
    private boolean isNewRtGroup(Cell cellA, Cell cellB) {
        if (cellA == null || cellB == null) return false;

        String valueA = getCellStringValue(cellA).trim();
        String valueB = getCellStringValue(cellB).trim();

        // Новая группа РТ: есть название РТ в столбце A и "УЗД" в столбце B
        boolean isRtFormat = valueA.matches("РТ-?\\d+.*");
        boolean isUzdType = valueB.contains("УЗД");

        return isRtFormat && isUzdType;
    }

    /**
     * Обрабатывает строку данных внутри группы РТ
     */
    private void processDataRow(FileData data, Cell cellB, Cell cellL) {
        String dataType = getCellStringValue(cellB).trim();
        Double value = getNumericValue(cellL);

        if (dataType.contains("УЗД")) {
            // УЗД днём/ночью - основное значение шума
            data.noiseLevel = value;
        } else if (dataType.contains("ПДУ")) {
            // ПДУ или ПДУ пом. - допустимый уровень
            data.pduValue = value;
        } else if (dataType.contains("превышение")) {
            // Превышение - текстовое значение (+/-)
            data.excessValue = value != null ? (value > 0 ? "+" : "-") : "";
        }
    }

    /**
     * Заполняет данные блока
     */
    private void fillBlockData(Row noiseRow, Row pduRow, Row excessRow, Map<String, FileData> fileData, List<String> rtNames) {
        for (int i = 0; i < rtNames.size(); i++) {
            String rtName = rtNames.get(i);
            FileData data = fileData.get(rtName);
            int colIndex = i + 1;

            if (data != null) {
                // УЗД данные
                if (data.noiseLevel != null) {
                    noiseRow.createCell(colIndex).setCellValue(data.noiseLevel);
                }

                // ПДУ значения
                if (data.pduValue != null) {
                    pduRow.createCell(colIndex).setCellValue(data.pduValue);
                }

                // Превышения
                if (data.excessValue != null) {
                    excessRow.createCell(colIndex).setCellValue(data.excessValue);
                }
            }
        }
    }

    /**
     * Применяет стили к блоку данных
     */
    private void applyBlockStyles(Workbook workbook, Sheet sheet, int startRow, int endRow, int numColumns) {
        CellStyle dataStyle = createDataStyle(workbook);

        for (int rowNum = startRow; rowNum <= endRow; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                for (int colNum = 0; colNum < numColumns; colNum++) {
                    Cell cell = row.getCell(colNum);
                    if (cell != null) {
                        cell.setCellStyle(dataStyle);
                    }
                }
            }
        }
    }

    /**
     * Создает стиль для заголовков
     */
    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return style;
    }

    /**
     * Создает стиль для данных
     */
    private CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    // Остальные методы без изменений (findAllSourceFiles, createOutputFolder, createOutputFile, etc.)
    private List<File> findAllSourceFiles(String rootPath) {
        try {
            List<File> sourceFiles = new ArrayList<>();
            FileType[] allTypes = FileType.values();

            Files.walk(Paths.get(rootPath))
                    .filter(Files::isRegularFile)
                    .filter(path -> {
                        String fileName = path.getFileName().toString();
                        String lowerFileName = fileName.toLowerCase();
                        return (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) &&
                                FileType.isSupportedFile(fileName) &&
                                !lowerFileName.contains("в записку"); // ← исключаем
                    })
                    .forEach(path -> sourceFiles.add(path.toFile()));

            log.info("Найдено файлов для сводной таблицы: {}", sourceFiles.size());
            return sourceFiles;

        } catch (IOException e) {
            log.error("Ошибка при поиске файлов: {}", e.getMessage());
            return Collections.emptyList();
        }
    }

    private File createOutputFolder(String rootPath) {
        File rootDir = new File(rootPath);
        String folderName = rootDir.getName() + "_Сводная таблица УЗД в расчетных точках, в дБА";
        File outputFolder = new File(rootDir, folderName);

        if (!outputFolder.exists()) {
            if (outputFolder.mkdirs()) {
                log.info("✅ Создана папка: {}", outputFolder.getAbsolutePath());
            } else {
                log.error("❌ Не удалось создать папку: {}", outputFolder.getAbsolutePath());
            }
        }

        return outputFolder;
    }

    private File createOutputFile(File outputFolder) {
        String fileName = outputFolder.getName() + ".xlsx";
        return new File(outputFolder, fileName);
    }

    private boolean isRtRow(Cell cellA, Cell cellB) {
        if (cellA == null || cellB == null) return false;
        String valueA = getCellStringValue(cellA).trim();
        String valueB = getCellStringValue(cellB).trim();
        boolean isRtFormat = valueA.matches("РТ-?\\d+.*");
        boolean isValidType = valueB.contains("УЗД") || valueB.contains("ПДУ") || valueB.contains("превышение");
        return isRtFormat && isValidType;
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        try {
            switch (cell.getCellType()) {
                case STRING: return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    } else {
                        double value = cell.getNumericCellValue();
                        return value == Math.floor(value) ? String.valueOf((int) value) : String.valueOf(value);
                    }
                case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    try { return cell.getStringCellValue(); }
                    catch (Exception e) {
                        try { return String.valueOf(cell.getNumericCellValue()); }
                        catch (Exception ex) { return cell.getCellFormula(); }
                    }
                default: return "";
            }
        } catch (Exception e) {
            return "";
        }
    }

    private Double getNumericValue(Cell cell) {
        if (cell == null || cell.getCellType() != CellType.NUMERIC) return null;
        try {
            return cell.getNumericCellValue();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Вспомогательный класс для хранения данных файла
     */
    private static class FileData {
        Double noiseLevel;
        Double pduValue;
        String excessValue;
    }

    // Метод для обратной совместимости
    public boolean createSummaryTable(String rootPath, boolean createSummaryTable) {
        return createTable(rootPath, createSummaryTable);
    }
}