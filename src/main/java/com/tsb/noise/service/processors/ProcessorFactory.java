package com.tsb.noise.service.processors;

import com.tsb.noise.model.FileType;

public class ProcessorFactory {

    public static BaseExcelProcessor createProcessor(FileType fileType) {
        if (fileType == null) {
            return new DayProcessor(); // По умолчанию
        }

        switch (fileType) {
            case TX_DAY:
            case OV_DAY:
            case POS_DAY:
                return new DayProcessor();

            case TX_NIGHT:
            case OV_NIGHT:
            case POS_NIGHT:
                return new NightProcessor();

            default:
                return new DayProcessor();
        }
    }

    public static BaseExcelProcessor createProcessor(String fileName) {
        FileType fileType = FileType.fromFileName(fileName);
        return createProcessor(fileType);
    }
}