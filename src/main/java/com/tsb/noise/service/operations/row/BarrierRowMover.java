package com.tsb.noise.service.operations.row;

import com.tsb.noise.service.operations.core.RowOperation;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

@Slf4j
public class BarrierRowMover implements RowOperation {

    private static final String TARGET_TEXT = "Звукоизоляция преградой";
    private static final int TARGET_COLUMN = 1; // Колонка B
    private static final int MOVE_OFFSET = 3; // На 3 строки выше

    @Override
    public int execute(Sheet sheet) {
        log.info("🔍 Поиск строк с '{}' для перемещения...", TARGET_TEXT);

        List<BarrierRowInfo> barrierRows = findTargetRows(sheet);

        if (barrierRows.isEmpty()) {
            log.info("❌ Строки с '{}' не найдены", TARGET_TEXT);
            return 0;
        }

        log.info("✅ Найдено строк: {}", barrierRows.size());
        barrierRows.sort(Comparator.comparingInt(BarrierRowInfo::getOriginalIndex).reversed());

        int movedCount = 0;

        for (BarrierRowInfo barrierRow : barrierRows) {
            try {
                if (moveBarrierRow(sheet, barrierRow)) {
                    movedCount++;
                }
            } catch (Exception e) {
                log.error("❌ Ошибка при перемещении строки {}: {}",
                        barrierRow.getOriginalIndex() + 1, e.getMessage(), e);
            }
        }

        log.info("🎯 Перемещено строк '{}': {}", TARGET_TEXT, movedCount);
        return movedCount;
    }

    @Override
    public String getOperationName() {
        return "Перемещение строк 'Звукоизоляция преградой'";
    }

    /**
     * Перемещает одну строку
     */
    private boolean moveBarrierRow(Sheet sheet, BarrierRowInfo barrierRow) {
        int sourceIndex = barrierRow.getOriginalIndex();
        int targetIndex = barrierRow.getTargetIndex();

        log.debug("🔄 Перемещение из {} в {}", sourceIndex + 1, targetIndex + 1);

        try {
            // Сохраняем данные исходной строки
            List<CellData> sourceRowData = saveRowData(barrierRow.getRow());

            // Вставляем пустую строку в целевую позицию
            insertEmptyRow(sheet, targetIndex);

            // Копируем данные в новую строку
            Row newRow = sheet.getRow(targetIndex);
            restoreRowData(newRow, sourceRowData);

            // Удаляем исходную строку (учитывая что мы вставили строку выше)
            removeRow(sheet, sourceIndex + 1);

            log.debug("✅ Успешно перемещено: исходная строка {} удалена, данные скопированы в {}",
                    sourceIndex + 1, targetIndex + 1);
            return true;

        } catch (Exception e) {
            log.error("❌ Критическая ошибка при перемещении строки из {} в {}: {}",
                    sourceIndex + 1, targetIndex + 1, e.getMessage(), e);
            return false;
        }
    }

    /**
     * Находит строки с целевым текстом
     */
    private List<BarrierRowInfo> findTargetRows(Sheet sheet) {
        List<BarrierRowInfo> barrierRows = new ArrayList<>();

        for (int rowIndex = 3; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(TARGET_COLUMN);
                if (cell != null) {
                    String cellValue = getCellStringValue(cell).trim();
                    if (TARGET_TEXT.equals(cellValue)) {
                        int targetIndex = rowIndex - MOVE_OFFSET;
                        if (targetIndex >= 3) { // Не выше шапки
                            barrierRows.add(new BarrierRowInfo(rowIndex, targetIndex, cellValue, row));
                        }
                    }
                }
            }
        }
        return barrierRows;
    }

    /**
     * Вспомогательные классы для хранения данных
     */
    private static class BarrierRowInfo {
        private final int originalIndex;
        private final int targetIndex;
        private final String cellValue;
        private final Row row;

        public BarrierRowInfo(int originalIndex, int targetIndex, String cellValue, Row row) {
            this.originalIndex = originalIndex;
            this.targetIndex = targetIndex;
            this.cellValue = cellValue;
            this.row = row;
        }

        public int getOriginalIndex() { return originalIndex; }
        public int getTargetIndex() { return targetIndex; }
        public String getCellValue() { return cellValue; }
        public Row getRow() { return row; }
    }

    private static class CellData {
        private final int columnIndex;
        private final String value;
        private final CellStyle style;

        public CellData(int columnIndex, String value, CellStyle style) {
            this.columnIndex = columnIndex;
            this.value = value;
            this.style = style;
        }
    }

    /**
     * Сохраняет данные строки
     */
    private List<CellData> saveRowData(Row row) {
        List<CellData> rowData = new ArrayList<>();
        for (int colIndex = 0; colIndex <= 13; colIndex++) {
            Cell cell = row.getCell(colIndex);
            if (cell != null) {
                String value = getCellStringValue(cell);
                rowData.add(new CellData(colIndex, value, cell.getCellStyle()));
            }
        }
        return rowData;
    }

    /**
     * Восстанавливает данные строки
     */
    private void restoreRowData(Row row, List<CellData> rowData) {
        for (CellData cellData : rowData) {
            Cell cell = row.createCell(cellData.columnIndex);
            cell.setCellValue(cellData.value);
            if (cellData.style != null) {
                cell.setCellStyle(cellData.style);
            }
        }
        // Сохраняем высоту строки
        if (!rowData.isEmpty()) {
            row.setHeightInPoints((short) (8.0 / 25.4 * 72)); // 8mm в points
        }
    }

    /**
     * Вставляет пустую строку
     */
    private void insertEmptyRow(Sheet sheet, int targetRowIndex) {
        if (targetRowIndex <= sheet.getLastRowNum()) {
            sheet.shiftRows(targetRowIndex, sheet.getLastRowNum(), 1, true, false);
        }

        Row newRow = sheet.createRow(targetRowIndex);
        newRow.setHeightInPoints((short) (8.0 / 25.4 * 72));

        for (int colIndex = 0; colIndex <= 13; colIndex++) {
            Cell cell = newRow.createCell(colIndex);
            cell.setCellValue("");
        }
    }

    /**
     * Удаляет строку
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