package com.bigcay.excel.excelfiller.template;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class ColumnTemplate extends AbstractTemplate {

    private String listName; 
    
    private String methodName; 
    
    public String getListName() {
        return listName;
    }

    public void setListName(String listName) {
        this.listName = listName;
    }

    public String getMethodName() {
        return methodName;
    }

    public void setMethodName(String methodName) {
        this.methodName = methodName;
    }
    
    public ColumnTemplate(int row, int col, String templateType, String templateCode, XSSFCellStyle templateStyle) {
        super(row, col, templateType, templateCode, templateStyle);
    }
}
