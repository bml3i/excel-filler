package com.bigcay.excel.excelfiller.style;

public class DefaultStyle {
    
    private Align align;
    
    private String datePattern;
    
    public DefaultStyle() {
        this.align = Align.RIGHT;
        this.datePattern = "m/d/yyyy";
    }
    
    public Align getAlign() {
        return align;
    }
    
    public void setAlign(Align align) {
        this.align = align;
    }
    
    public String getDatePattern() {
        return datePattern;
    }
    
    public void setDatePattern(String datePattern) {
        this.datePattern = datePattern;
    }
    
}
