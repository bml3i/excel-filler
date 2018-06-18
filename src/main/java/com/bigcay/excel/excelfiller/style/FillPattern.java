package com.bigcay.excel.excelfiller.style;

import org.apache.poi.ss.usermodel.FillPatternType;

public enum FillPattern {
    
    NO_FILL(FillPatternType.NO_FILL),
    
    SOLID_FOREGROUND(FillPatternType.SOLID_FOREGROUND),
    
    FINE_DOTS(FillPatternType.FINE_DOTS),
    
    ALT_BARS(FillPatternType.ALT_BARS),
    
    SPARSE_DOTS(FillPatternType.SPARSE_DOTS),
    
    THICK_HORZ_BANDS(FillPatternType.THICK_HORZ_BANDS),
    
    THICK_VERT_BANDS(FillPatternType.THICK_VERT_BANDS),
    
    THICK_BACKWARD_DIAG(FillPatternType.THICK_BACKWARD_DIAG),
    
    THICK_FORWARD_DIAG(FillPatternType.THICK_FORWARD_DIAG),
    
    BIG_SPOTS(FillPatternType.BIG_SPOTS),
    
    BRICKS(FillPatternType.BRICKS),
    
    THIN_HORZ_BANDS(FillPatternType.THIN_HORZ_BANDS),
    
    THIN_VERT_BANDS(FillPatternType.THIN_VERT_BANDS),
    
    THIN_BACKWARD_DIAG(FillPatternType.THIN_BACKWARD_DIAG),
    
    THIN_FORWARD_DIAG(FillPatternType.THIN_FORWARD_DIAG),
    
    SQUARES(FillPatternType.SQUARES),
    
    DIAMONDS(FillPatternType.DIAMONDS),
    
    LESS_DOTS(FillPatternType.LESS_DOTS),
    
    LEAST_DOTS(FillPatternType.LEAST_DOTS);
    
    private FillPatternType fillPattern;
    
    private FillPattern(FillPatternType fillPattern) {
        this.fillPattern = fillPattern;
    }
    
    public FillPatternType getFillPattern() {
        return fillPattern;
    }
    
}
