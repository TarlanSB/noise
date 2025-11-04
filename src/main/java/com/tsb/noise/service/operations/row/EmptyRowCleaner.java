package com.tsb.noise.service.operations.row;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

@Slf4j
public class EmptyRowCleaner {

    public int removeEmptyRows(Sheet sheet) {
        int removedCount = 0;

        // Идем с конца чтобы индексы не сбивались
        for (int rowIndex = sheet.getLastRowNum(); rowIndex >= 0; rowIndex--) {
            Row row = sheet.getRow(rowIndex);
            if (isRowEmpty(row) && rowIndex > 2) { // Не удаляем шапку (строки 0-2)
                removeRow(sheet, rowIndex);
                removedCount++;
            }
        }

        log.info("Удалено пустых строк: {}", removedCount);
        return removedCount;
    }

    private boolean isRowEmpty(Row row) {
        if (row == null) return true;

        for (int cellIndex = 0; cellIndex <= 12; cellIndex++) { // A-M (после скрытия C)
            if (cellIndex == 2) continue; // Пропускаем скрытую колонку C
            Cell cell = row.getCell(cellIndex);
            if (cell != null) {
                String cellValue = getCellStringValue(cell);
                if (cellValue != null && !cellValue.trim().isEmpty()) {
                    return false;
                }
            }
        }
        return true;
    }

    private void removeRow(Sheet sheet, int rowIndex) {
        if (rowIndex < sheet.getLastRowNum()) {
            sheet.shiftRows(rowIndex + 1, sheet.getLastRowNum(), -1);
        } else {
            sheet.removeRow(sheet.getRow(rowIndex));
        }
    }

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