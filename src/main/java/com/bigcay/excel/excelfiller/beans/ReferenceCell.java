package com.bigcay.excel.excelfiller.beans;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

public class ReferenceCell {

	public ReferenceCell(int row, int col) {
		this.row = row;
		this.col = col;
	}
	
	public ReferenceCell(int row, int col, HSSFCellStyle cellStyle) {
		this(row, col);
		this.cellStyle = cellStyle;
	}

	private int row;

	private int col;

	private HSSFCellStyle cellStyle;

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

	public HSSFCellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(HSSFCellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

}
