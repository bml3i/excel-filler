package com.bigcay.excel.excelfiller.element;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.util.WorkbookUtil;

import com.bigcay.excel.excelfiller.ExcelContext;

public class SheetElement extends AbstractElement {
    
    private HSSFSheet sheet;
    
    private int sheetIndex;
    
    public SheetElement(ExcelContext excelContext) {
        super(excelContext);
    }
    
    public SheetElement(HSSFSheet sheet, ExcelContext excelContext) {
        super(excelContext);
        this.sheet = sheet;
        sheetIndex = workbook.getSheetIndex(this.sheet);
    }

    public int getSheetIndex() {
        return sheetIndex;
    }
    
	public SheetElement sheetName(String sheetName) {
		String safeSheetName = WorkbookUtil.createSafeSheetName(sheetName);
		workbook.setSheetName(sheetIndex, safeSheetName);
		return this;
	}
}
