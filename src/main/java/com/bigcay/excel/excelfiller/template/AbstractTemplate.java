package com.bigcay.excel.excelfiller.template;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public abstract class AbstractTemplate {

	protected int row;

	protected int col;

	protected String templateType;

	protected XSSFCellStyle templateStyle;

	protected String templateCode;

	private boolean autoRowHeight;

	private int rowCharNumber;

	private float unitRowHeightInPoint;

	public AbstractTemplate(int row, int col, String templateType,
			String templateCode) {
		this.row = row;
		this.col = col;
		this.templateType = templateType;
		this.templateCode = templateCode;
	}
	
	public AbstractTemplate(int row, int col, String templateType,
			String templateCode, XSSFCellStyle templateStyle) {
		this(row, col, templateType,templateCode);
		this.templateStyle = templateStyle;
	}

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

	public String getTemplateType() {
		return templateType;
	}

	public void setTemplateType(String templateType) {
		this.templateType = templateType;
	}

	public XSSFCellStyle getTemplateStyle() {
		return templateStyle;
	}

	public void setTemplateStyle(XSSFCellStyle templateStyle) {
		this.templateStyle = templateStyle;
	}

	public String getTemplateCode() {
		return templateCode;
	}

	public void setTemplateCode(String templateCode) {
		this.templateCode = templateCode;
	}

	public boolean isAutoRowHeight() {
		return autoRowHeight;
	}

	public void setAutoRowHeight(boolean autoRowHeight) {
		this.autoRowHeight = autoRowHeight;
	}

	public int getRowCharNumber() {
		return rowCharNumber;
	}

	public void setRowCharNumber(int rowCharNumber) {
		this.rowCharNumber = rowCharNumber;
	}

	public float getUnitRowHeightInPoint() {
		return unitRowHeightInPoint;
	}

	public void setUnitRowHeightInPoint(float unitRowHeightInPoint) {
		this.unitRowHeightInPoint = unitRowHeightInPoint;
	}

}
