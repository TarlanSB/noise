package com.tsb.noise.service.operations.table;

import com.tsb.noise.service.operations.core.StyleApplier;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

@Slf4j
public class DataCopier {

    private final StyleApplier styleApplier;

    public DataCopier(StyleApplier styleApplier) {
        this.styleApplier = styleApplier;
    }

    public void copyDataFromSource(Sheet sourceSheet, Sheet targetSheet, double rowHeightMm) {
        int targetRowIndex = 3; // Начинаем с четвертой строки

        for (int sourceRowIndex = 1; sourceRowIndex <= sourceSheet.getLastRowNum(); sourceRowIndex++) {
            Row sourceRow = sourceSheet.getRow(sourceRowIndex);
            if (sourceRow == null) continue;

            // Создаем строку с фиксированной высотой
            Row targetRow = createRowWithFixedHeight(targetSheet, targetRowIndex, rowHeightMm);

            // Копируем данные включая колонку C
            copyRowWithAllColumns(sourceRow, targetRow);

            targetRowIndex++;
        }

        log.info("Скопировано {} строк данных с высотой 8мм", targetRowIndex - 3);
    }

    private void copyRowWithAllColumns(Row sourceRow, Row targetRow) {
        for (int sourceColIndex = 0; sourceColIndex < sourceRow.getLastCellNum(); sourceColIndex++) {
            Cell sourceCell = sourceRow.getCell(sourceColIndex);
            if (sourceCell == null) continue;

            // Копируем в те же колонки (включая C)
            Cell targetCell = targetRow.createCell(sourceColIndex);
            copyCellValue(sourceCell, targetCell);
            styleApplier.applyCellStyleWithFont(targetCell);
        }
    }

    private void copyCellValue(Cell sourceCell, Cell targetCell) {
        switch (sourceCell.getCellType()) {
            case STRING:
                targetCell.setCellValue(sourceCell.getStringCellValue());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(sourceCell)) {
                    targetCell.setCellValue(sourceCell.getDateCellValue());
                } else {
                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                targetCell.setCellValue(sourceCell.getBooleanCellValue());
                break;
            case FORMULA:
                targetCell.setCellFormula(sourceCell.getCellFormula());
                break;
            case BLANK:
                targetCell.setBlank();
                break;
            default:
                targetCell.setCellValue("");
        }
    }

    private Row createRowWithFixedHeight(Sheet sheet, int rowIndex, double rowHeightMm) {
        Row row = sheet.createRow(rowIndex);
        row.setHeightInPoints(mmToPoints(rowHeightMm));
        return row;
    }

    private short mmToPoints(double mm) {
        return (short) (mm / 25.4 * 72);
    }
}