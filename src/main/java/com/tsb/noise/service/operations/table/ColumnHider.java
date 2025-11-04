package com.tsb.noise.service.operations.table;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

@Slf4j
public class ColumnHider {

    /**
     * Скрывает столбец C где находится "31,5" во второй строке
     */
    public void hideColumnC(Sheet sheet) {
        log.info("🔍 Поиск столбца C для скрытия...");

        try {
            Row headerRow2 = sheet.getRow(1);
            if (headerRow2 == null) {
                log.warn("❌ Вторая строка шапки не найдена, невозможно определить столбец для скрытия");
                return;
            }

            int columnToHide = findColumnWithValue(headerRow2, "31,5");

            if (columnToHide != -1) {
                sheet.setColumnHidden(columnToHide, true);
                log.info("✅ Столбец C (индекс {}) с '31,5' скрыт", columnToHide);
            } else {
                log.warn("⚠️ Столбец с '31,5' не найден во второй строке шапки");
            }

        } catch (Exception e) {
            log.error("❌ Ошибка при скрытии столбца C: {}", e.getMessage(), e);
        }
    }

    /**
     * Находит столбец с указанным значением в строке
     */
    private int findColumnWithValue(Row row, String targetValue) {
        for (int colIndex = 0; colIndex <= row.getLastCellNum(); colIndex++) {
            Cell cell = row.getCell(colIndex);
            if (cell != null) {
                String cellValue = getCellStringValue(cell).trim();
                if (targetValue.equals(cellValue)) {
                    log.debug("Найден столбец с '{}': индекс {}", targetValue, colIndex);
                    return colIndex;
                }
            }
        }
        return -1;
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
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            default:
                return "";
        }
    }
}