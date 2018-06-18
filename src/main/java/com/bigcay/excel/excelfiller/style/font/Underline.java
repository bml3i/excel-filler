package com.bigcay.excel.excelfiller.style.font;

import org.apache.poi.xssf.usermodel.XSSFFont;

public enum Underline {
	
	NONE(XSSFFont.U_NONE),

	SINGLE(XSSFFont.U_SINGLE),

	DOUBLE(XSSFFont.U_DOUBLE),

	SINGLE_ACCOUNTING(XSSFFont.U_SINGLE_ACCOUNTING),

	DOUBLE_ACCOUNTING(XSSFFont.U_DOUBLE_ACCOUNTING);

	private byte line;

	private Underline(byte line) {
		this.line = line;
	}

	public byte getLine() {
		return line;
	}

	public static Underline instance(byte line) {
		for (Underline underline : Underline.values()) {
			if (underline.getLine() == line) {
				return underline;
			}
		}
		return Underline.NONE;
	}

}
