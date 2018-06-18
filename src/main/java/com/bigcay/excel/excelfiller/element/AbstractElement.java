package com.bigcay.excel.excelfiller.element;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.bigcay.excel.excelfiller.ExcelContext;
import com.bigcay.excel.excelfiller.util.ExcelUtil;

public abstract class AbstractElement {

	protected XSSFWorkbook workbook;
	protected XSSFSheet workingSheet;
	protected XSSFCellStyle tempCellStyle;
	protected XSSFFont tempFont;
	protected ExcelContext excelContext;

	public AbstractElement(ExcelContext excelContext) {
		this.excelContext = excelContext;
		this.workbook = excelContext.getWorkbook();
		this.workingSheet = excelContext.getWorkingSheet();
		this.tempCellStyle = excelContext.getTempCellStyle();
		this.tempFont = excelContext.getTempFont();
	}

	protected XSSFRow getRow(int row) {
		return ExcelUtil.getXSSFRow(this.workingSheet, row);
	}

	protected XSSFCell getCell(int row, int col) {
		return ExcelUtil.getXSSFCell(this.workingSheet, row, col);
	}

	protected XSSFCell getCell(XSSFRow row, int col) {
		return ExcelUtil.getXSSFCell(row, col);
	}
}
