package com.tsb.noise.service.utils;

import com.tsb.noise.model.RtData;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
public class RtDataProcessor {

    private static final Pattern COORDINATES_PATTERN = Pattern.compile(":[^:]*:([^,]*)");
    private static final Pattern ELEVATION_PATTERN = Pattern.compile("(-?\\d+\\.?\\d*)");

    // Константа для высоты строк 8мм
    private static final double ROW_HEIGHT_MM = 8.0;

    /**
     * Конвертирует мм в points для высоты строк
     */
    private short mmToPoints(double mm) {
        return (short) (mm / 25.4 * 72);
    }

    /**
     * Создает строку с фиксированной высотой 8мм
     */
    private Row createRowWithFixedHeight(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        row.setHeightInPoints(mmToPoints(ROW_HEIGHT_MM));
        return row;
    }

    public void processRtData(Sheet sourceSheet, Sheet targetSheet) {
        try {
            List<RtData> rtDataList = findRtData(sourceSheet);
            log.info("Найдено РТ для обработки: {}", rtDataList.size());

            if (rtDataList.isEmpty()) {
                log.info("РТ для обработки не найдены");
                return;
            }

            // Сортируем по убыванию индекса строки (идем с конца)
            rtDataList.sort(Comparator.comparingInt(RtData::getRowIndex).reversed());

            // Обрабатываем все найденные РТ с конца
            for (RtData rtData : rtDataList) {
                createRtHeaderBeforeData(targetSheet, rtData);
            }

            log.info("Обработка РТ завершена. Добавлено описаний: {}", rtDataList.size());

        } catch (Exception e) {
            log.error("Ошибка при обработке данных РТ: {}", e.getMessage(), e);
            throw new RuntimeException("Ошибка обработки данных РТ", e);
        }
    }

    /**
     * Создает заголовок РТ непосредственно перед его данными с высотой 8мм
     */
    private void createRtHeaderBeforeData(Sheet sheet, RtData rtData) {
        try {
            int targetRowIndex = rtData.getRowIndex() + 1;

            if (targetRowIndex < 1 || targetRowIndex > sheet.getLastRowNum() + 1) {
                log.warn("Некорректный индекс строки для РТ {}: {}. LastRowNum: {}",
                        rtData.getName(), targetRowIndex, sheet.getLastRowNum());
                return;
            }

            // Сдвигаем строки вниз начиная с целевой позиции
            int lastRowNum = sheet.getLastRowNum();
            if (targetRowIndex <= lastRowNum) {
                sheet.shiftRows(targetRowIndex, lastRowNum, 1, true, false);
            }

            // Создаем строку с фиксированной высотой 8мм
            Row headerRow = createRowWithFixedHeight(sheet, targetRowIndex);

            // Создаем ячейку B с объединенным текстом
            Cell headerCell = headerRow.createCell(1);
            String headerText = buildHeaderText(rtData);
            headerCell.setCellValue(headerText);

            // Объединяем ячейки B-M
            CellRangeAddress mergedRegion = new CellRangeAddress(targetRowIndex, targetRowIndex, 1, 12);
            sheet.addMergedRegion(mergedRegion);

            // Применяем стиль
            applyRtHeaderStyle(headerCell);

            log.debug("✅ Создано описание РТ '{}' в строке {} с высотой 8мм", headerText, targetRowIndex + 1);

        } catch (Exception e) {
            log.error("Ошибка при создании описания РТ {}: {}", rtData.getName(), e.getMessage(), e);
        }
    }

    /**
     * Применяет стиль к заголовку РТ с ВЫРАВНИВАНИЕМ ПО ЦЕНТРУ БЕЗ автопереноса
     */
    private void applyRtHeaderStyle(Cell cell) {
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            CellStyle style = workbook.createCellStyle();

            // ВЫРАВНИВАНИЕ ПО ЦЕНТРУ И ПО ВЕРТИКАЛИ
            style.setAlignment(HorizontalAlignment.CENTER);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setWrapText(false); // УБРАН автоперенос

            // Шрифт Arial Narrow
            Font font = workbook.createFont();
            font.setFontName("Arial Narrow");
            font.setBold(true);
            font.setFontHeightInPoints((short) 11);
            style.setFont(font);

            // Настраиваем границы
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);

            // Заливка
            style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            cell.setCellStyle(style);
        } catch (Exception e) {
            log.warn("Не удалось применить стиль к заголовку РТ: {}", e.getMessage());
        }
    }

    /**
     * Находит данные РТ в исходном листе
     */
    private List<RtData> findRtData(Sheet sheet) {
        List<RtData> rtDataList = new ArrayList<>();

        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell cellA = row.getCell(0);
            Cell cellB = row.getCell(1);
            Cell cellO = row.getCell(14);
            Cell cellN = row.getCell(13);

            if (isRtRow(cellA, cellB)) {
                RtData rtData = extractRtData(cellA, cellO, cellN, rowIndex);
                if (rtData != null) {
                    rtDataList.add(rtData);
                    log.debug("Найдено РТ: {} в строке {}", rtData.getName(), rowIndex + 1);
                }
            }
        }
        return rtDataList;
    }

    /**
     * Проверяет, является ли строка строкой РТ
     * Теперь учитывает как "УЗД днём", так и "УЗД ночью"
     */
    private boolean isRtRow(Cell cellA, Cell cellB) {
        if (cellA == null || cellB == null) return false;

        String valueA = getCellStringValue(cellA).trim();
        String valueB = getCellStringValue(cellB).trim();

        // Проверяем формат названия РТ и наличие УЗД днём/ночью
        boolean isRtFormat = valueA.matches("РТ-?\\d+.*");
        boolean isUzdDayOrNight = "УЗД днём".equals(valueB) || "УЗД ночью".equals(valueB);

        return isRtFormat && isUzdDayOrNight;
    }

    /**
     * Извлекает данные РТ из ячеек
     */
    private RtData extractRtData(Cell cellA, Cell cellO, Cell cellN, int rowIndex) {
        try {
            String name = getCellStringValue(cellA).trim();
            String description = cellO != null ? getCellStringValue(cellO).trim() : "";
            String coordinates = cellN != null ? getCellStringValue(cellN).trim() : "";
            Double elevation = extractElevation(coordinates);

            // Логируем для отладки
            log.debug("Извлечение РТ: name={}, description={}, coordinates={}, elevation={}",
                    name, description, coordinates, elevation);

            return new RtData(name, description, coordinates, elevation, rowIndex);
        } catch (Exception e) {
            log.warn("Не удалось извлечь данные РТ из строки {}: {}", rowIndex + 1, e.getMessage());
            return null;
        }
    }

    /**
     * Извлекает отметку высоты из координат
     */
    private Double extractElevation(String coordinates) {
        if (coordinates == null || coordinates.isEmpty()) return null;
        try {
            Matcher coordMatcher = COORDINATES_PATTERN.matcher(coordinates);
            if (coordMatcher.find()) {
                String elevationStr = coordMatcher.group(1).trim();
                Matcher elevationMatcher = ELEVATION_PATTERN.matcher(elevationStr);
                if (elevationMatcher.find()) {
                    return Double.parseDouble(elevationMatcher.group(1));
                }
            }
        } catch (Exception e) {
            log.debug("Не удалось извлечь высоту из координат: {}", coordinates);
        }
        return null;
    }

    /**
     * Строит текст заголовка РТ
     */
    private String buildHeaderText(RtData rtData) {
        StringBuilder sb = new StringBuilder();
        sb.append(rtData.getName());
        if (!rtData.getDescription().isEmpty()) {
            sb.append(" ").append(rtData.getDescription());
        }
        if (rtData.getElevation() != null) {
            sb.append(String.format(", отметка %+.3f", rtData.getElevation()));
        }

        String result = sb.toString();
        log.debug("Построен заголовок РТ: {}", result);
        return result;
    }

    /**
     * Получает строковое значение ячейки
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
}