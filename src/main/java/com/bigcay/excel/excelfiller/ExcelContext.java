package com.bigcay.excel.excelfiller;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.bigcay.excel.excelfiller.style.DefaultStyle;

public class ExcelContext {
    
    private XSSFWorkbook workbook;
    
    private Map<String, Object> excelData = new HashMap<String, Object>();
    
    private XSSFCellStyle tempCellStyle;
    
    private XSSFFont tempFont;
    
    private Map<Integer, XSSFCellStyle> stylePool = new HashMap<Integer, XSSFCellStyle>();
    
    private Map<Integer, XSSFFont> fontPool = new HashMap<Integer, XSSFFont>();
    
    private DefaultStyle defaultStyle;
    
    private XSSFSheet workingSheet;
    
    private int workingSheetIndex = 0;
    
    protected ExcelContext(XSSFWorkbook workbook) {
        this.workbook = workbook;
        
        int numStyle = workbook.getNumCellStyles();
        for (int i = 0; i < numStyle; i++) {
            XSSFCellStyle style = workbook.getCellStyleAt(i);
            if (style != tempCellStyle) {
                stylePool.put(style.hashCode() - style.getIndex(), style);
            }
        }
        
        short numFont = workbook.getNumberOfFonts();
        for (short i = 0; i < numFont; i++) {
            XSSFFont font = workbook.getFontAt(i);
            if (font != tempFont) {
                fontPool.put(font.hashCode() - font.getIndex(), font);
            }
        }
    }

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public XSSFCellStyle getTempCellStyle() {
        return tempCellStyle;
    }

    public void setTempCellStyle(XSSFCellStyle tempCellStyle) {
        this.tempCellStyle = tempCellStyle;
    }

    public XSSFFont getTempFont() {
        return tempFont;
    }

    public void setTempFont(XSSFFont tempFont) {
        this.tempFont = tempFont;
    }

    public Map<Integer, XSSFCellStyle> getStylePool() {
        return stylePool;
    }

    public void setStylePool(Map<Integer, XSSFCellStyle> stylePool) {
        this.stylePool = stylePool;
    }

    public Map<Integer, XSSFFont> getFontPool() {
        return fontPool;
    }

    public void setFontPool(Map<Integer, XSSFFont> fontPool) {
        this.fontPool = fontPool;
    }

    public DefaultStyle getDefaultStyle() {
        return defaultStyle;
    }

    public void setDefaultStyle(DefaultStyle defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    public XSSFSheet getWorkingSheet() {
        return workingSheet;
    }

    public void setWorkingSheet(XSSFSheet workingSheet) {
        this.workingSheet = workingSheet;
    }

    public int getWorkingSheetIndex() {
        return workingSheetIndex;
    }

    public void setWorkingSheetIndex(int workingSheetIndex) {
        this.workingSheetIndex = workingSheetIndex;
    }

    public Map<String, Object> getExcelData() {
        return excelData;
    }

    public void setExcelData(Map<String, Object> excelData) {
        this.excelData = excelData;
    };
    
}
