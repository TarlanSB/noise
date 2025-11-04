package com.tsb.noise.service.operations.row;

import com.tsb.noise.service.operations.core.RowOperation;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

@Slf4j
public class SoundIsolationRemover implements RowOperation {

    private static final String TARGET_TEXT = "Требуемая звукоизоляция";
    private static final int TARGET_COLUMN = 1; // Колонка B

    @Override
    public int execute(Sheet sheet) {
        log.info("🔍 Поиск строк с '{}' для удаления...", TARGET_TEXT);

        List<Integer> rowsToRemove = findTargetRows(sheet);

        if (rowsToRemove.isEmpty()) {
            log.info("❌ Строки с '{}' не найдены", TARGET_TEXT);
            return 0;
        }

        log.info("✅ Найдено строк для удаления: {}", rowsToRemove.size());

        // Удаляем с конца чтобы индексы не сбивались
        rowsToRemove.sort((a, b) -> b - a);
        int removedCount = 0;

        for (int rowIndex : rowsToRemove) {
            try {
                removeRow(sheet, rowIndex);
                removedCount++;
                log.debug("✅ Удалена строка с '{}' в строке {}", TARGET_TEXT, rowIndex + 1);
            } catch (Exception e) {
                log.error("❌ Ошибка при удалении строки {}: {}", rowIndex + 1, e.getMessage(), e);
            }
        }

        log.info("🎯 Удалено строк '{}': {}", TARGET_TEXT, removedCount);
        return removedCount;
    }

    @Override
    public String getOperationName() {
        return "Удаление строк 'Требуемая звукоизоляция'";
    }

    /**
     * Находит строки с целевым текстом
     */
    private List<Integer> findTargetRows(Sheet sheet) {
        List<Integer> rowsToRemove = new ArrayList<>();

        for (int rowIndex = sheet.getLastRowNum(); rowIndex >= 0; rowIndex--) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(TARGET_COLUMN);
                if (cell != null) {
                    String cellValue = getCellStringValue(cell).trim();
                    if (TARGET_TEXT.equals(cellValue)) {
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