package com.bigcay.excel.excelfiller.template;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

public class MatrixTemplate extends AbstractTemplate {

	public MatrixTemplate(int row, int col, String templateType, String templateCode, HSSFCellStyle cellStyle) {
		super(row, col, templateType, templateCode, cellStyle);
	}

}
