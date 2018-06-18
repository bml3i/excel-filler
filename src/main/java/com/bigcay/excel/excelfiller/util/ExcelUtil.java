package com.bigcay.excel.excelfiller.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class ExcelUtil {
    
    public static HSSFSheet getHSSFSheet(HSSFWorkbook workbook, int index) {
        if (index > workbook.getNumberOfSheets() - 1) {
            workbook.createSheet();
            return workbook.getSheetAt(workbook.getNumberOfSheets() - 1);
        } else {
            return workbook.getSheetAt(index);
        }
    }
    
    public static HSSFCell getHSSFCell(HSSFSheet sheet, int row, int col) {
        HSSFRow r = getHSSFRow(sheet, row);
        return getHSSFCell(r, col);
    }
    
    public static HSSFRow getHSSFRow(HSSFSheet sheet, int row) {
        HSSFRow r = sheet.getRow(row);
        if (r == null) {
            r = sheet.createRow(row);
        }
        return r;
    }
    
    public static HSSFCell getHSSFCell(HSSFRow row, int col) {
        HSSFCell c = row.getCell(col);
        if (c == null) {
            c = row.createCell(col);
        }
        return c;
    }
}
