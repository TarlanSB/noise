package com.tsb.noise.service.operations.export;

import com.tsb.noise.model.FileType;
import com.tsb.noise.service.operations.core.SheetLayoutManager;
import com.tsb.noise.service.operations.core.StyleApplier;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;


/**
 * Создатель сводной таблицы расчетных точек
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
     * Создает сводную таблицу РТ
     */
    @Override
    public boolean createTable(String rootPath, boolean createSummaryTable) {
        if (!createSummaryTable) {
            log.info("Создание сводной таблицы отключено");
            return false;
        }

        log.info("🚀 Начало создания сводной таблицы РТ...");

        try {
            // Находим все файлы для обработки
            List<File> sourceFiles = findAllSourceFiles(rootPath);
            if (sourceFiles.isEmpty()) {
                log.error("❌ Не найдены файлы для создания сводной таблицы");
                return false;
            }

            log.info("✅ Найдено файлов для обработки: {}", sourceFiles.size());

            // Извлекаем данные из всех файлов
            List<SummaryData> summaryDataList = extractSummaryDataFromFiles(sourceFiles);
            if (summaryDataList.isEmpty()) {
                log.warn("⚠️ Не найдены данные для сводной таблицы");
                return false;
            }

            log.info("✅ Извлечено данных РТ: {}", summaryDataList.size());

            // Создаем папку и файл
            File outputFolder = createOutputFolder(rootPath);
            File outputFile = createOutputFile(outputFolder);

            // Создаем сводную таблицу
            return createSummaryWorkbook(summaryDataList, outputFile);

        } catch (Exception e) {
            log.error("❌ Ошибка при создании сводной таблицы: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * Метод для обратной совместимости
     */
    public boolean createSummaryTable(String rootPath, boolean createSummaryTable) {
        return createTable(rootPath, createSummaryTable);
    }

    /**
     * Находит все файлы в директории
     */
    private List<File> findAllSourceFiles(String rootPath) {
        File rootDir = new File(rootPath);
        if (!rootDir.exists() || !rootDir.isDirectory()) {
            log.error("Корневая папка не существует: {}", rootPath);
            return Collections.emptyList();
        }

        List<File> sourceFiles = new ArrayList<>();
        FileType[] allTypes = FileType.values();

        for (FileType fileType : allTypes) {
            File[] files = rootDir.listFiles((dir, name) ->
                    name.contains(fileType.getInputPattern()) && (name.endsWith(".xlsx") || name.endsWith(".xls"))
            );

            if (files != null) {
                Collections.addAll(sourceFiles, files);
            }
        }

        log.debug("Найдено файлов: {}", sourceFiles.size());
        return sourceFiles;
    }

    /**
     * Извлекает сводные данные из всех файлов
     */
    private List<SummaryData> extractSummaryDataFromFiles(List<File> sourceFiles) {
        Map<String, SummaryData> summaryDataMap = new HashMap<>();

        for (File sourceFile : sourceFiles) {
            try (FileInputStream fis = new FileInputStream(sourceFile);
                 Workbook workbook = WorkbookFactory.create(fis)) {

                FileType fileType = FileType.fromFileName(sourceFile.getName());
                if (fileType == null) continue;

                Sheet sheet = workbook.getSheet("ЛИСТ2");
                if (sheet == null) continue;

                extractDataFromSheet(sheet, fileType, summaryDataMap, sourceFile.getName());

            } catch (IOException e) {
                log.error("Ошибка при чтении файла {}: {}", sourceFile.getName(), e.getMessage(), e);
            }
        }

        List<SummaryData> result = new ArrayList<>(summaryDataMap.values());
        result.sort(Comparator.comparing(SummaryData::getRtName));
        return result;
    }

    /**
     * Извлекает данные из листа
     */
    private void extractDataFromSheet(Sheet sheet, FileType fileType, Map<String, SummaryData> summaryDataMap, String fileName) {
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell cellA = row.getCell(0); // Наименование РТ
            Cell cellB = row.getCell(1); // Тип (УЗД днём/ночью, ПДУ и т.д.)
            Cell cellL = row.getCell(11); // Lэкв, дБА (колонка L)
            Cell cellN = row.getCell(13); // Координаты с отметкой

            if (isRtRow(cellA, cellB)) {
                String rtName = getCellStringValue(cellA).trim();
                String dataType = getCellStringValue(cellB).trim();
                Double leqvValue = getNumericValue(cellL);
                String coordinates = getCellStringValue(cellN);

                SummaryData summaryData = summaryDataMap.computeIfAbsent(rtName,
                        k -> new SummaryData(rtName, extractElevation(coordinates)));

                // Заполняем данные в зависимости от типа
                fillSummaryData(summaryData, fileType, dataType, leqvValue, fileName);

                log.debug("Извлечены данные для РТ {}: тип={}, Lэкв={}", rtName, dataType, leqvValue);
            }
        }
    }

    /**
     * Заполняет сводные данные
     */
    private void fillSummaryData(SummaryData summaryData, FileType fileType, String dataType, Double leqvValue, String fileName) {
        String timeSuffix = fileName.contains("ночь") ? " (ночь)" : " (день)";
        String fileTypeName = fileType.getDisplayName() + timeSuffix;

        switch (dataType) {
            case "УЗД днём":
            case "УЗД ночью":
                summaryData.getLeqvValues().put(fileTypeName, leqvValue);
                break;
            case "ПДУ":
            case "ПДУ пом.":
                summaryData.getPduValues().put(fileTypeName, leqvValue);
                break;
            case "превышение":
            case "превышение пом.":
                summaryData.getExcessValues().put(fileTypeName, leqvValue);
                break;
        }
    }

    /**
     * Проверяет, является ли строка строкой РТ
     */
    private boolean isRtRow(Cell cellA, Cell cellB) {
        if (cellA == null || cellB == null) return false;

        String valueA = getCellStringValue(cellA).trim();
        String valueB = getCellStringValue(cellB).trim();

        boolean isRtFormat = valueA.matches("РТ-?\\d+.*");
        boolean isValidType = "УЗД днём".equals(valueB) || "УЗД ночью".equals(valueB) ||
                "ПДУ".equals(valueB) || "ПДУ пом.".equals(valueB) ||
                "превышение".equals(valueB) || "превышение пом.".equals(valueB);

        return isRtFormat && isValidType;
    }

    /**
     * Извлекает отметку высоты из координат
     */
    private Double extractElevation(String coordinates) {
        if (coordinates == null || coordinates.isEmpty()) return null;
        try {
            // Пример координат: "x:123.456 y:789.012 z:15.678"
            String[] parts = coordinates.split(" ");
            for (String part : parts) {
                if (part.startsWith("z:") || part.contains(":")) {
                    String[] keyValue = part.split(":");
                    if (keyValue.length == 2) {
                        return Double.parseDouble(keyValue[1].trim());
                    }
                }
            }
        } catch (Exception e) {
            log.debug("Не удалось извлечь высоту из координат: {}", coordinates);
        }
        return null;
    }

    /**
     * Создает папку для вывода
     */
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

    /**
     * Создает файл для вывода
     */
    private File createOutputFile(File outputFolder) {
        String fileName = outputFolder.getName() + ".xlsx";
        return new File(outputFolder, fileName);
    }

    /**
     * Создает рабочую книгу со сводной таблицей
     */
    private boolean createSummaryWorkbook(List<SummaryData> summaryDataList, File outputFile) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Сводная таблица УЗД");

            // Настраиваем layout
            setupSheetLayout(sheet);

            // Создаем шапку таблицы
            createTableHeader(workbook, sheet, summaryDataList);

            // Заполняем данными
            fillTableData(workbook, sheet, summaryDataList);

            // Применяем стили
            styleApplier.applyTableBorders(sheet);

            // Сохраняем файл
            outputFile.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            log.info("✅ Успешно создана сводная таблица: {}", outputFile.getAbsolutePath());
            return true;

        } catch (IOException e) {
            log.error("❌ Ошибка при создании сводной таблицы: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * Настраивает layout листа
     */
    private void setupSheetLayout(Sheet sheet) {
        // Устанавливаем ширину колонок (18см = ~18*4.5*256 = 20736 units)
        int columnWidthUnits = (int) (18.0 * 4.5 * 256);

        sheet.setColumnWidth(0, columnWidthUnits / 6); // Колонка A - 3см
        sheet.setColumnWidth(1, columnWidthUnits / 6); // Колонка B - 3см

        // Динамическая ширина для остальных колонок
        for (int i = 2; i < 20; i++) {
            sheet.setColumnWidth(i, columnWidthUnits / 12); // По 1.5см
        }

        // Высота строк
        sheet.setDefaultRowHeightInPoints(20);
    }

    /**
     * Создает шапку таблицы
     */
    private void createTableHeader(Workbook workbook, Sheet sheet, List<SummaryData> summaryDataList) {
        // Первая строка - основные заголовки
        Row headerRow1 = sheet.createRow(0);
        headerRow1.setHeightInPoints(25);

        // Вторая строка - подзаголовки
        Row headerRow2 = sheet.createRow(1);
        headerRow2.setHeightInPoints(20);

        // Заполняем заголовки
        createHeaderCells(workbook, headerRow1, headerRow2, summaryDataList);

        // Объединяем ячейки для основных заголовков
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 1));
    }

    /**
     * Создает ячейки заголовков
     */
    private void createHeaderCells(Workbook workbook, Row headerRow1, Row headerRow2, List<SummaryData> summaryDataList) {
        // Основные заголовки
        createHeaderCell(workbook, headerRow1, 0, "Расчетные точки");
        createHeaderCell(workbook, headerRow1, 1, "Отметка, м");

        // Получаем все уникальные типы данных
        Set<String> allDataTypes = getAllDataTypes(summaryDataList);
        int colIndex = 2;

        for (String dataType : allDataTypes) {
            createHeaderCell(workbook, headerRow1, colIndex, dataType);
            createHeaderCell(workbook, headerRow2, colIndex, "Lэкв, дБА");
            colIndex++;
        }
    }

    /**
     * Получает все уникальные типы данных
     */
    private Set<String> getAllDataTypes(List<SummaryData> summaryDataList) {
        Set<String> dataTypes = new TreeSet<>();
        for (SummaryData data : summaryDataList) {
            dataTypes.addAll(data.getLeqvValues().keySet());
            dataTypes.addAll(data.getPduValues().keySet());
            dataTypes.addAll(data.getExcessValues().keySet());
        }
        return dataTypes;
    }

    /**
     * Заполняет таблицу данными
     */
    private void fillTableData(Workbook workbook, Sheet sheet, List<SummaryData> summaryDataList) {
        Set<String> allDataTypes = getAllDataTypes(summaryDataList);
        List<String> sortedDataTypes = new ArrayList<>(allDataTypes);
        Collections.sort(sortedDataTypes);

        for (int i = 0; i < summaryDataList.size(); i++) {
            SummaryData data = summaryDataList.get(i);
            Row row = sheet.createRow(i + 2); // +2 потому что две строки заголовков

            // Наименование РТ
            Cell cellA = row.createCell(0);
            cellA.setCellValue(data.getRtName());

            // Отметка высоты
            Cell cellB = row.createCell(1);
            if (data.getElevation() != null) {
                cellB.setCellValue(data.getElevation());
            } else {
                cellB.setCellValue("-");
            }

            // Данные по типам
            int colIndex = 2;
            for (String dataType : sortedDataTypes) {
                Cell cell = row.createCell(colIndex);

                Double leqvValue = data.getLeqvValues().get(dataType);
                Double pduValue = data.getPduValues().get(dataType);
                Double excessValue = data.getExcessValues().get(dataType);

                // Приоритет: Lэкв > ПДУ > превышение
                if (leqvValue != null) {
                    cell.setCellValue(leqvValue);
                } else if (pduValue != null) {
                    cell.setCellValue(pduValue);
                } else if (excessValue != null) {
                    cell.setCellValue(excessValue);
                } else {
                    cell.setCellValue("-");
                }

                styleApplier.applyCellStyleWithFont(cell);
                colIndex++;
            }

            // Применяем стиль к основным ячейкам
            styleApplier.applyCellStyleWithFont(cellA);
            styleApplier.applyCellStyleWithFont(cellB);
        }

        log.debug("Заполнено {} строк данными", summaryDataList.size());
    }

    /**
     * Создает ячейку заголовка
     */
    private void createHeaderCell(Workbook workbook, Row headerRow, int colIndex, String value) {
        Cell cell = headerRow.createCell(colIndex);
        cell.setCellValue(value);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);

        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 10);
        headerStyle.setFont(font);

        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        cell.setCellStyle(headerStyle);
    }

    /**
     * Вспомогательный метод для получения строкового значения ячейки
     */
    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        try {
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        return cell.getDateCellValue().toString();
                    } else {
                        double value = cell.getNumericCellValue();
                        if (value == Math.floor(value) && !Double.isInfinite(value)) {
                            return String.valueOf((int) value);
                        } else {
                            return String.valueOf(value);
                        }
                    }
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception e) {
                        try {
                            return String.valueOf(cell.getNumericCellValue());
                        } catch (Exception ex) {
                            return cell.getCellFormula();
                        }
                    }
                default:
                    return "";
            }
        } catch (Exception e) {
            return "";
        }
    }

    /**
     * Вспомогательный метод для получения числового значения ячейки
     */
    private Double getNumericValue(Cell cell) {
        if (cell == null || cell.getCellType() != CellType.NUMERIC) return null;
        try {
            return cell.getNumericCellValue();
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * Класс для хранения сводных данных
     */
    private static class SummaryData {
        private final String rtName;
        private final Double elevation;
        private final Map<String, Double> leqvValues = new HashMap<>(); // Lэкв по типам
        private final Map<String, Double> pduValues = new HashMap<>();  // ПДУ по типам
        private final Map<String, Double> excessValues = new HashMap<>(); // Превышения по типам

        public SummaryData(String rtName, Double elevation) {
            this.rtName = rtName;
            this.elevation = elevation;
        }

        public String getRtName() { return rtName; }
        public Double getElevation() { return elevation; }
        public Map<String, Double> getLeqvValues() { return leqvValues; }
        public Map<String, Double> getPduValues() { return pduValues; }
        public Map<String, Double> getExcessValues() { return excessValues; }
    }
}