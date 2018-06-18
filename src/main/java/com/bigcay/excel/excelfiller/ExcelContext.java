package com.bigcay.excel.excelfiller;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.bigcay.excel.excelfiller.style.DefaultStyle;

public class ExcelContext {
    
    private HSSFWorkbook workbook;
    
    private Map<String, Object> excelData = new HashMap<String, Object>();
    
    private HSSFCellStyle tempCellStyle;
    
    private HSSFFont tempFont;
    
    private Map<Integer, HSSFCellStyle> stylePool = new HashMap<Integer, HSSFCellStyle>();
    
    private Map<Integer, HSSFFont> fontPool = new HashMap<Integer, HSSFFont>();
    
    private DefaultStyle defaultStyle;
    
    private HSSFSheet workingSheet;
    
    private int workingSheetIndex = 0;
    
    protected ExcelContext(HSSFWorkbook workbook) {
        this.workbook = workbook;
        
        int numStyle = workbook.getNumCellStyles();
        for (int i = 0; i < numStyle; i++) {
            HSSFCellStyle style = workbook.getCellStyleAt(i);
            if (style != tempCellStyle) {
                stylePool.put(style.hashCode() - style.getIndex(), style);
            }
        }
        
        short numFont = workbook.getNumberOfFonts();
        for (short i = 0; i < numFont; i++) {
            HSSFFont font = workbook.getFontAt(i);
            if (font != tempFont) {
                fontPool.put(font.hashCode() - font.getIndex(), font);
            }
        }
    }

    public HSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(HSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public HSSFCellStyle getTempCellStyle() {
        return tempCellStyle;
    }

    public void setTempCellStyle(HSSFCellStyle tempCellStyle) {
        this.tempCellStyle = tempCellStyle;
    }

    public HSSFFont getTempFont() {
        return tempFont;
    }

    public void setTempFont(HSSFFont tempFont) {
        this.tempFont = tempFont;
    }

    public Map<Integer, HSSFCellStyle> getStylePool() {
        return stylePool;
    }

    public void setStylePool(Map<Integer, HSSFCellStyle> stylePool) {
        this.stylePool = stylePool;
    }

    public Map<Integer, HSSFFont> getFontPool() {
        return fontPool;
    }

    public void setFontPool(Map<Integer, HSSFFont> fontPool) {
        this.fontPool = fontPool;
    }

    public DefaultStyle getDefaultStyle() {
        return defaultStyle;
    }

    public void setDefaultStyle(DefaultStyle defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    public HSSFSheet getWorkingSheet() {
        return workingSheet;
    }

    public void setWorkingSheet(HSSFSheet workingSheet) {
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
