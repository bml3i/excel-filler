package com.bigcay.excel.excelfiller.template;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public class MatrixTemplate extends AbstractTemplate {

	public MatrixTemplate(int row, int col, String templateType, String templateCode, XSSFCellStyle cellStyle) {
		super(row, col, templateType, templateCode, cellStyle);
	}

}
