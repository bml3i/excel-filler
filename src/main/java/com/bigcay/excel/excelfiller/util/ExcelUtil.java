package com.bigcay.excel.excelfiller.util;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {
    
    public static XSSFSheet getXSSFSheet(XSSFWorkbook workbook, int index) {
        if (index > workbook.getNumberOfSheets() - 1) {
            workbook.createSheet();
            return workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
        } else {
            return workbook.getSheetAt(index);
        }
    }
    
    public static XSSFCell getXSSFCell(XSSFSheet sheet, int row, int col) {
        XSSFRow r = getXSSFRow(sheet, row);
        return getXSSFCell(r, col);
    }
    
    public static XSSFRow getXSSFRow(XSSFSheet sheet, int row) {
        XSSFRow r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        return r;
    }
    
    public static XSSFCell getXSSFCell(XSSFRow row, int col) {
        XSSFCell c = row.getCell(col);
        if (c == null) {
            c = row.createCell(col);
        }
        return c;
    }
}
