package com.tsb.noise.service.operations;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

@Slf4j
public class RowRemover {

    /**
     * Удаляет строки с "Требуемая звукоизоляция" из колонки B
     */
    public void removeSoundIsolationRows(Sheet sheet) {
        log.info("🔍 Поиск строк с 'Требуемая звукоизоляция' для удаления...");

        List<Integer> rowsToRemove = findSoundIsolationRows(sheet);

        if (rowsToRemove.isEmpty()) {
            log.info("❌ Строки с 'Требуемая звукоизоляция' не найдены");
            return;
        }

        log.info("✅ Найдено строк для удаления: {}", rowsToRemove.size());

        // Удаляем с конца чтобы индексы не сбивались
        rowsToRemove.sort((a, b) -> b - a);
        int removedCount = 0;

        for (int rowIndex : rowsToRemove) {
            try {
                removeRow(sheet, rowIndex);
                removedCount++;
                log.debug("✅ Удалена строка с 'Требуемая звукоизоляция' в строке {}", rowIndex + 1);
            } catch (Exception e) {
                log.error("❌ Ошибка при удалении строки {}: {}", rowIndex + 1, e.getMessage(), e);
            }
        }

        log.info("🎯 Удалено строк 'Требуемая звукоизоляция': {}", removedCount);
    }

    /**
     * Находит строки с "Требуемая звукоизоляция"
     */
    private List<Integer> findSoundIsolationRows(Sheet sheet) {
        List<Integer> rowsToRemove = new ArrayList<>();

        for (int rowIndex = sheet.getLastRowNum(); rowIndex >= 0; rowIndex--) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cellB = row.getCell(1); // Колонка B
                if (cellB != null) {
                    String cellValue = getCellStringValue(cellB).trim();
                    if ("Требуемая звукоизоляция".equals(cellValue)) {
                        rowsToRemove.add(rowIndex);
                    }
                }
            }
        }

        return rowsToRemove;
    }

    /**
     * Физически удаляет строку используя shiftRows
     */
    private void removeRow(Sheet sheet, int rowIndex) {
        if (rowIndex >= 0 && rowIndex <= sheet.getLastRowNum()) {
            if (rowIndex < sheet.getLastRowNum()) {
                sheet.shiftRows(rowIndex + 1, sheet.getLastRowNum(), -1);
            } else {
                // Если это последняя строка
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }
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