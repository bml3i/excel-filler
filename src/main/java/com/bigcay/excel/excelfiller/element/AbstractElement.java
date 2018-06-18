package com.bigcay.excel.excelfiller.element;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.bigcay.excel.excelfiller.ExcelContext;
import com.bigcay.excel.excelfiller.util.ExcelUtil;

public abstract class AbstractElement {

	protected HSSFWorkbook workbook;
	protected HSSFSheet workingSheet;
	protected HSSFCellStyle tempCellStyle;
	protected HSSFFont tempFont;
	protected ExcelContext excelContext;

	public AbstractElement(ExcelContext excelContext) {
		this.excelContext = excelContext;
		this.workbook = excelContext.getWorkbook();
		this.workingSheet = excelContext.getWorkingSheet();
		this.tempCellStyle = excelContext.getTempCellStyle();
		this.tempFont = excelContext.getTempFont();
	}

	protected HSSFRow getRow(int row) {
		return ExcelUtil.getHSSFRow(this.workingSheet, row);
	}

	protected HSSFCell getCell(int row, int col) {
		return ExcelUtil.getHSSFCell(this.workingSheet, row, col);
	}

	protected HSSFCell getCell(HSSFRow row, int col) {
		return ExcelUtil.getHSSFCell(row, col);
	}
}
