package com.tsb.noise.service.processors;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

@Slf4j
public class OvDataProcessor {

    private static final String[] TARGET_TEXTS = {"ПДУ", "ПДУ пом."};
    private static final int TARGET_COLUMN = 1; // Колонка B
    private static final String CORRECTION_SUFFIX = " c учётом поправки -5 дБ";

    /**
     * Обрабатывает данные для файлов ОВ - добавляет поправку к ПДУ
     */
    public void processOvData(Sheet sheet) {
        log.info("🔍 Поиск ячеек с ПДУ для добавления поправки -5 дБ...");

        int processedCells = 0;

        for (int rowIndex = 3; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(TARGET_COLUMN);
                if (cell != null && isTargetCell(cell)) {
                    try {
                        addCorrectionToCell(cell);
                        processedCells++;
                        log.debug("✅ Добавлена поправка к ПДУ в строке {}", rowIndex + 1);
                    } catch (Exception e) {
                        log.error("❌ Ошибка при добавлении поправки к ячейке в строке {}: {}",
                                rowIndex + 1, e.getMessage(), e);
                    }
                }
            }
        }

        log.info("🎯 Добавлена поправка -5 дБ к {} ячейкам с ПДУ", processedCells);
    }

    /**
     * Проверяет, является ли ячейка целевой (содержит ПДУ или ПДУ пом.)
     */
    private boolean isTargetCell(Cell cell) {
        String cellValue = getCellStringValue(cell).trim();
        for (String targetText : TARGET_TEXTS) {
            if (targetText.equals(cellValue)) {
                return true;
            }
        }
        return false;
    }

    /**
     * Добавляет поправку к ячейке
     */
    private void addCorrectionToCell(Cell cell) {
        String originalValue = getCellStringValue(cell).trim();
        String newValue = originalValue + CORRECTION_SUFFIX;

        cell.setCellValue(newValue);

        // Сохраняем стиль ячейки
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(cell.getCellStyle());

            // Устанавливаем перенос текста для длинного текста
            newStyle.setWrapText(true);
            cell.setCellStyle(newStyle);

            log.debug("🔄 Обновлена ячейка: '{}' -> '{}'", originalValue, newValue);
        } catch (Exception e) {
            log.warn("⚠️ Не удалось сохранить стиль ячейки: {}", e.getMessage());
        }
    }

    /**
     * Вспомогательный метод для получения строкового значения ячейки
     */
    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";

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
    }
}