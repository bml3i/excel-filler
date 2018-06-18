package com.bigcay.excel.excelfiller.style;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

public enum Align {
    GENERAL(HorizontalAlignment.GENERAL),
    LEFT(HorizontalAlignment.LEFT),
    CENTER(HorizontalAlignment.CENTER),
    RIGHT(HorizontalAlignment.RIGHT),
    FILL(HorizontalAlignment.FILL),
    JUSTIFY(HorizontalAlignment.JUSTIFY),
    CENTER_SELECTION(HorizontalAlignment.CENTER_SELECTION);
    
    private HorizontalAlignment alignment;
    
    private Align(HorizontalAlignment alignment) {
        this.alignment = alignment;
    }
    
    public HorizontalAlignment getAlignment() {
        return alignment;
    }
}
