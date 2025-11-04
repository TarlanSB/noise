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
import java.util.*;

@Slf4j
public class RtListCreator {

    private final StyleApplier styleApplier;
    private final SheetLayoutManager layoutManager;

    public RtListCreator() {
        this.styleApplier = new StyleApplier();
        this.layoutManager = new SheetLayoutManager();
    }

    /**
     * Создает таблицу "Перечень расчетных точек"
     */
    public boolean createRtListTable(String rootPath, boolean createRtList) {
        if (!createRtList) {
            log.info("Создание перечня РТ отключено");
            return false;
        }

        log.info("🚀 Начало создания перечня расчетных точек...");
        log.info("📁 Поиск файлов в директории: {}", rootPath);

        try {
            // Находим подходящий файл для извлечения данных
            File sourceFile = findSourceFileForRtList(rootPath);
            if (sourceFile == null) {
                log.error("❌ Не найден подходящий файл для создания перечня РТ");
                log.info("🔍 Проверьте наличие файлов с паттернами: 'УЗД в РТ ОВ', 'УЗД в РТ ТХ', 'УЗД в РТ ПОС'");
                return false;
            }

            log.info("✅ Используется файл для извлечения данных: {}", sourceFile.getName());

            // Извлекаем данные РТ
            List<RtData> rtDataList = extractRtDataFromFile(sourceFile);
            if (rtDataList.isEmpty()) {
                log.warn("⚠️ В файле не найдены данные РТ");
                log.info("🔍 Проверьте наличие строк РТ на листе 'ЛИСТ2' в столбцах A, N, O");
                return false;
            }

            log.info("✅ Извлечено уникальных РТ: {}", rtDataList.size());

            // Создаем папку и файл
            File outputFolder = createOutputFolder(rootPath);
            File outputFile = createOutputFile(outputFolder);

            // Создаем таблицу
            boolean result = createRtListWorkbook(rtDataList, outputFile);

            if (result) {
                log.info("🎉 Перечень расчетных точек успешно создан: {}", outputFile.getAbsolutePath());
            } else {
                log.error("❌ Не удалось создать файл перечня РТ");
            }

            return result;

        } catch (Exception e) {
            log.error("❌ Ошибка при создании перечня РТ: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * Находит подходящий файл для извлечения данных РТ по приоритету: ОВ -> ТХ -> ПОС
     */
    private File findSourceFileForRtList(String rootPath) {
        File rootDir = new File(rootPath);
        if (!rootDir.exists() || !rootDir.isDirectory()) {
            log.error("❌ Корневая папка не существует: {}", rootPath);
            return null;
        }

        log.info("🔍 Сканирование директории: {}", rootDir.getAbsolutePath());

        // Получаем все Excel файлы в директории
        File[] allFiles = rootDir.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xls")
        );

        if (allFiles == null || allFiles.length == 0) {
            log.warn("⚠️ В директории не найдено Excel файлов");
            return null;
        }

        log.info("📊 Найдено Excel файлов: {}", allFiles.length);

        // Логируем все найденные файлы для отладки
        for (File file : allFiles) {
            log.debug("📄 Найден файл: {}", file.getName());
        }

        // Приоритет поиска: ОВ -> ТХ -> ПОС
        String[] priorityPatterns = {
                "УЗД в РТ ОВ",  // Первый приоритет
                "УЗД в РТ ТХ",  // Второй приоритет
                "УЗД в РТ ПОС"  // Третий приоритет
        };

        for (String pattern : priorityPatterns) {
            File foundFile = findFileByPattern(allFiles, pattern);
            if (foundFile != null) {
                log.info("✅ Найден файл по паттерну '{}': {}", pattern, foundFile.getName());
                return foundFile;
            } else {
                log.debug("❌ Файл с паттерном '{}' не найден", pattern);
            }
        }

        log.warn("⚠️ Не найден ни один подходящий файл для создания перечня РТ");
        log.info("🔍 Доступные файлы в директории:");
        for (File file : allFiles) {
            log.info("   - {}", file.getName());
        }
        return null;
    }

    /**
     * Ищет файл по паттерну в названии (частичное совпадение)
     */
    private File findFileByPattern(File[] allFiles, String pattern) {
        for (File file : allFiles) {
            if (file.getName().contains(pattern)) {
                log.debug("🎯 Найден файл содержащий '{}': {}", pattern, file.getName());

                // Дополнительная проверка - файл должен быть читаемым
                if (!file.canRead()) {
                    log.warn("⚠️ Файл недоступен для чтения: {}", file.getName());
                    continue;
                }

                // Проверяем, что файл содержит нужный лист
                if (hasRequiredSheet(file)) {
                    return file;
                } else {
                    log.warn("⚠️ Файл не содержит лист 'ЛИСТ2': {}", file.getName());
                }
            }
        }
        return null;
    }

    /**
     * Проверяет, содержит ли файл необходимый лист
     */
    private boolean hasRequiredSheet(File file) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("ЛИСТ2");
            boolean hasSheet = sheet != null;

            if (hasSheet) {
                log.debug("✅ Файл содержит лист 'ЛИСТ2': {}", file.getName());
            } else {
                log.debug("❌ Файл не содержит лист 'ЛИСТ2': {}", file.getName());
                // Логируем доступные листы для отладки
                log.debug("📋 Доступные листы в файле {}:", file.getName());
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    log.debug("   - {}", workbook.getSheetName(i));
                }
            }

            return hasSheet;
        } catch (Exception e) {
            log.warn("⚠️ Ошибка при проверке файла {}: {}", file.getName(), e.getMessage());
            return false;
        }
    }

    /**
     * Извлекает данные РТ из файла
     */
    private List<RtData> extractRtDataFromFile(File sourceFile) {
        List<RtData> rtDataList = new ArrayList<>();
        Set<String> uniqueRtNames = new HashSet<>();

        log.info("📖 Извлечение данных РТ из файла: {}", sourceFile.getName());

        try (FileInputStream fis = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet("ЛИСТ2");
            if (sheet == null) {
                log.error("❌ Лист 'ЛИСТ2' не найден в файле: {}", sourceFile.getName());
                return rtDataList;
            }

            log.info("📊 Обработка листа 'ЛИСТ2', строк: {}", sheet.getLastRowNum());

            int rtCount = 0;
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;

                Cell cellA = row.getCell(0); // Наименование РТ (столбец A)
                Cell cellN = row.getCell(13); // Координаты (столбец N)
                Cell cellO = row.getCell(14); // Описание (столбец O)

                if (isValidRtRow(cellA)) {
                    RtData rtData = extractRtData(cellA, cellN, cellO);
                    if (rtData != null && uniqueRtNames.add(rtData.getName())) {
                        rtDataList.add(rtData);
                        rtCount++;
                        log.debug("📍 Извлечен РТ: {} (строка {})", rtData.getName(), rowIndex + 1);
                    }
                }
            }

            log.info("✅ Извлечено уникальных РТ: {}", rtCount);

        } catch (IOException e) {
            log.error("❌ Ошибка при чтении файла {}: {}", sourceFile.getName(), e.getMessage(), e);
        } catch (Exception e) {
            log.error("❌ Неожиданная ошибка при обработке файла {}: {}", sourceFile.getName(), e.getMessage(), e);
        }

        // Сортируем по имени РТ
        rtDataList.sort(Comparator.comparing(RtData::getName));
        return rtDataList;
    }

    /**
     * Проверяет, является ли строка валидной строкой РТ
     */
    private boolean isValidRtRow(Cell cellA) {
        if (cellA == null) {
            return false;
        }

        String valueA = getCellStringValue(cellA).trim();

        // Проверяем формат названия РТ (РТ-1, РТ-2, РТ-10, РТ-15 и т.д.)
        boolean isRtFormat = valueA.matches("РТ-?\\d+.*");

        if (isRtFormat) {
            log.trace("✅ Валидная строка РТ: {}", valueA);
        } else {
            log.trace("❌ Невалидная строка РТ: {}", valueA);
        }

        return isRtFormat;
    }

    /**
     * Извлекает данные РТ
     */
    private RtData extractRtData(Cell cellA, Cell cellN, Cell cellO) {
        try {
            String name = getCellStringValue(cellA).trim();
            String coordinates = cellN != null ? getCellStringValue(cellN).trim() : "";
            String description = cellO != null ? getCellStringValue(cellO).trim() : "";

            // Проверяем, что имя РТ не пустое
            if (name.isEmpty()) {
                log.warn("⚠️ Пустое имя РТ в ячейке");
                return null;
            }

            RtData rtData = new RtData(name, coordinates, description);
            log.trace("📝 Данные РТ: name='{}', coordinates='{}', description='{}'",
                    name, coordinates, description);

            return rtData;
        } catch (Exception e) {
            log.warn("⚠️ Не удалось извлечь данные РТ: {}", e.getMessage());
            return null;
        }
    }

    /**
     * Создает папку для вывода
     */
    private File createOutputFolder(String rootPath) {
        File rootDir = new File(rootPath);
        String folderName = rootDir.getName() + "_Перечень РТ";
        File outputFolder = new File(rootDir, folderName);

        if (!outputFolder.exists()) {
            if (outputFolder.mkdirs()) {
                log.info("✅ Создана папка: {}", outputFolder.getAbsolutePath());
            } else {
                log.error("❌ Не удалось создать папку: {}", outputFolder.getAbsolutePath());
            }
        } else {
            log.info("✅ Папка уже существует: {}", outputFolder.getAbsolutePath());
        }

        return outputFolder;
    }

    /**
     * Создает файл для вывода
     */
    private File createOutputFile(File outputFolder) {
        String fileName = outputFolder.getName() + ".xlsx";
        File outputFile = new File(outputFolder, fileName);
        log.info("💾 Выходной файл: {}", outputFile.getAbsolutePath());
        return outputFile;
    }

    /**
     * Создает рабочую книгу с перечнем РТ
     */
    private boolean createRtListWorkbook(List<RtData> rtDataList, File outputFile) {
        log.info("🛠️ Создание Excel файла с {} РТ", rtDataList.size());

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Перечень РТ");

            // Настраиваем layout с шириной 18см
            setupSheetLayout(sheet);

            // Создаем шапку таблицы
            createTableHeader(workbook, sheet);

            // Заполняем данными
            fillTableData(workbook, sheet, rtDataList);

            // Применяем стили
            styleApplier.applyTableBorders(sheet);

            // Сохраняем файл
            outputFile.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            log.info("✅ Успешно создан файл перечня РТ: {}", outputFile.getAbsolutePath());
            return true;

        } catch (IOException e) {
            log.error("❌ Ошибка при создании файла перечня РТ: {}", e.getMessage(), e);
            return false;
        } catch (Exception e) {
            log.error("❌ Неожиданная ошибка при создании файла: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * Настраивает layout листа с шириной 18см
     */
    private void setupSheetLayout(Sheet sheet) {
        // 18см = ~18 * 4.5 * 256 = 20736 units
        int totalWidthUnits = (int) (18.0 * 4.5 * 256);

        // Распределяем ширину колонок (A: 6см, B: 6см, C: 6см)
        int columnWidth = totalWidthUnits / 3;

        sheet.setColumnWidth(0, columnWidth); // Колонка A - 6см
        sheet.setColumnWidth(1, columnWidth); // Колонка B - 6см
        sheet.setColumnWidth(2, columnWidth); // Колонка C - 6см

        // Высота строк
        sheet.setDefaultRowHeightInPoints(20);

        log.debug("📐 Настроен layout таблицы: ширина 18см, 3 колонки по 6см");
    }

    /**
     * Создает шапку таблицы
     */
    private void createTableHeader(Workbook workbook, Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        headerRow.setHeightInPoints(25);

        // Создаем ячейки шапки
        Cell cellA = headerRow.createCell(0);
        cellA.setCellValue("Наименование РТ");

        Cell cellB = headerRow.createCell(1);
        cellB.setCellValue("Координаты РТ x:y:z");

        Cell cellC = headerRow.createCell(2);
        cellC.setCellValue("Описание РТ");

        // Применяем стиль шапки
        applyHeaderStyle(workbook, headerRow);

        log.debug("📋 Создана шапка таблицы");
    }

    /**
     * Применяет стиль к строке заголовка
     */
    private void applyHeaderStyle(Workbook workbook, Row headerRow) {
        CellStyle headerStyle = workbook.createCellStyle();

        // Выравнивание по центру
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);

        // Шрифт
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        headerStyle.setFont(font);

        // Заливка
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Границы
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        // Применяем стиль ко всем ячейкам заголовка
        for (Cell cell : headerRow) {
            cell.setCellStyle(headerStyle);
        }
    }

    /**
     * Заполняет таблицу данными
     */
    private void fillTableData(Workbook workbook, Sheet sheet, List<RtData> rtDataList) {
        CellStyle dataStyle = createDataStyle(workbook);

        for (int i = 0; i < rtDataList.size(); i++) {
            RtData rtData = rtDataList.get(i);
            Row row = sheet.createRow(i + 1); // +1 потому что шапка в строке 0

            // Наименование РТ
            Cell cellA = row.createCell(0);
            cellA.setCellValue(rtData.getName());
            cellA.setCellStyle(dataStyle);

            // Координаты
            Cell cellB = row.createCell(1);
            cellB.setCellValue(rtData.getCoordinates());
            cellB.setCellStyle(dataStyle);

            // Описание
            Cell cellC = row.createCell(2);
            cellC.setCellValue(rtData.getDescription());
            cellC.setCellStyle(dataStyle);
        }

        log.debug("📊 Заполнено {} строк данными РТ", rtDataList.size());
    }

    /**
     * Создает стиль для данных
     */
    private CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // Выравнивание
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        // Шрифт
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Arial Narrow");
        style.setFont(font);

        // Границы
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
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
     * Внутренний класс для хранения данных РТ
     */
    private static class RtData {
        private final String name;
        private final String coordinates;
        private final String description;

        public RtData(String name, String coordinates, String description) {
            this.name = name;
            this.coordinates = coordinates;
            this.description = description;
        }

        public String getName() { return name; }
        public String getCoordinates() { return coordinates; }
        public String getDescription() { return description; }
    }
}