package com.tsb.noise.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class RtData {
    private String name;
    private String description;
    private String coordinates;
    private Double elevation;
    private int rowIndex;
}