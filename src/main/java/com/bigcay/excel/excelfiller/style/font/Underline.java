package com.bigcay.excel.excelfiller.style.font;

import org.apache.poi.hssf.usermodel.HSSFFont;

public enum Underline {
	
	NONE(HSSFFont.U_NONE),

	SINGLE(HSSFFont.U_SINGLE),

	DOUBLE(HSSFFont.U_DOUBLE),

	SINGLE_ACCOUNTING(HSSFFont.U_SINGLE_ACCOUNTING),

	DOUBLE_ACCOUNTING(HSSFFont.U_DOUBLE_ACCOUNTING);

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
