package com.tsb.noise.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class ProcessConfig {
    private boolean removeSoundIsolation;
    private boolean moveSoundIsolation;
    private FileType fileType;

    public static ProcessConfig defaultConfig() {
        return new ProcessConfig(false, false, null);
    }
}