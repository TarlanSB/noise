package com.tsb.noise.service.operations.table;

import com.tsb.noise.model.FrequencyBand;
import com.tsb.noise.service.operations.core.StyleApplier;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

@Slf4j
public class TableHeaderCreator {

    private static final short FONT_HEIGHT = 10;
    private static final String FONT_NAME = "Arial Narrow";

    private final StyleApplier styleApplier;

    public TableHeaderCreator(StyleApplier styleApplier) {
        this.styleApplier = styleApplier;
    }

    public void createTableHeader(Sheet sheet, double rowHeightMm) {
        // Создаем строки с фиксированной высотой
        Row headerRow1 = createRowWithFixedHeight(sheet, 0, rowHeightMm);
        Row headerRow2 = createRowWithFixedHeight(sheet, 1, rowHeightMm);

        // Ячейка B1-B2 объединенная "Наименование" (колонка 1)
        Cell cellB1 = headerRow1.createCell(1);
        cellB1.setCellValue("Наименование");

        // Создаем все колонки включая C с "31,5"
        Cell cellC1 = headerRow1.createCell(2);
        cellC1.setCellValue("Уровни звукового давления, дБ, в октавных полосах частот, Гц");

        // Ячейки L1 и M1 для Lэкв и Lмакс
        Cell cellL1 = headerRow1.createCell(11);
        cellL1.setCellValue("Lэкв, дБА");

        Cell cellM1 = headerRow1.createCell(12);
        cellM1.setCellValue("Lмакс, дБА");

        // Объединение ячеек
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 2, 10)); // C1-K1
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 1, 1));  // B1-B2
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 11, 11)); // L1-L2
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 12, 12)); // M1-M2

        // Заполняем частотные полосы
        int colIndex = 2;
        for (FrequencyBand band : FrequencyBand.values()) {
            if (band == FrequencyBand.LEKV || band == FrequencyBand.LMAX) {
                continue;
            }
            Cell cell = headerRow2.createCell(colIndex);
            cell.setCellValue(band.getDisplayName());
            colIndex++;
        }

        styleApplier.applyHeaderStyle(sheet.getWorkbook(), headerRow1, headerRow2);
        log.debug("Создана шапка таблицы со всеми колонками включая C с '31,5'");
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