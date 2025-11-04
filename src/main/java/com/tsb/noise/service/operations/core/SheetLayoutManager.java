package com.tsb.noise.service.operations.core;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;

@Slf4j
public class SheetLayoutManager {

    private static final double ROW_HEIGHT_MM = 8.0;
    private static final double COLUMN_WIDTH_B_CM = 3.0;
    private static final double COLUMN_WIDTH_OTHER_CM = 1.5;

    public void setupSheetLayout(Sheet sheet) {
        // Устанавливаем высоту строки 8мм для ВСЕХ строк
        short rowHeightInPoints = mmToPoints(ROW_HEIGHT_MM);
        sheet.setDefaultRowHeight(rowHeightInPoints);

        // Устанавливаем ширину колонок
        sheet.setColumnWidth(0, cmToUnits(1.5));  // Колонка A - 1.5см
        sheet.setColumnWidth(1, cmToUnits(COLUMN_WIDTH_B_CM)); // Колонка B - 3см

        // Устанавливаем ширину для всех остальных колонок (C-N)
        for (int i = 2; i <= 13; i++) {
            sheet.setColumnWidth(i, cmToUnits(COLUMN_WIDTH_OTHER_CM));
        }

        log.debug("Настроен layout листа: высота строк {}pt для всех строк", rowHeightInPoints);
    }

    private short mmToPoints(double mm) {
        return (short) (mm / 25.4 * 72);
    }

    private int cmToUnits(double cm) {
        return (int) (cm * 4.5 * 256);
    }
}