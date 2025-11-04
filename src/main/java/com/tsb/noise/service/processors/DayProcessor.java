package com.tsb.noise.service.processors;

import com.tsb.noise.model.ProcessConfig;
import com.tsb.noise.service.utils.RtDataProcessor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@Slf4j
public class DayProcessor extends BaseExcelProcessor {
    private final RtDataProcessor rtDataProcessor;

    public DayProcessor() {
        super();
        this.rtDataProcessor = new RtDataProcessor();
    }

    @Override
    protected Sheet getSourceSheet(Workbook workbook) {
        return workbook.getSheet("ЛИСТ2");
    }

    @Override
    protected void processSpecificData(Sheet sourceSheet, Sheet targetSheet, ProcessConfig config) {
        log.info("Обработка дневных данных РТ...");
        rtDataProcessor.processRtData(sourceSheet, targetSheet);
    }
}