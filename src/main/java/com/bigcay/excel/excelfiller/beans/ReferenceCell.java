package com.bigcay.excel.excelfiller.beans;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class ReferenceCell {

	public ReferenceCell(int row, int col) {
		this.row = row;
		this.col = col;
	}
	
	public ReferenceCell(int row, int col, XSSFCellStyle cellStyle) {
		this(row, col);
		this.cellStyle = cellStyle;
	}

	private int row;

	private int col;

	private XSSFCellStyle cellStyle;

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public int getCol() {
		return col;
	}

	public void setCol(int col) {
		this.col = col;
	}

	public XSSFCellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(XSSFCellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

}
