package com.tsb.noise.model;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

@Getter
@RequiredArgsConstructor
public enum FrequencyBand {
    BAND_31_5("31,5"),
    BAND_63("63"),
    BAND_125("125"),
    BAND_250("250"),
    BAND_500("500"),
    BAND_1000("1000"),
    BAND_2000("2000"),
    BAND_4000("4000"),
    BAND_8000("8000"),
    LEKV("Lэкв, дБА"),
    LMAX("Lмакс, дБА");

    private final String displayName;
}