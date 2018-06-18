package com.bigcay.excel.excelfiller.style.font;

import org.apache.poi.xssf.usermodel.XSSFFont;

import com.bigcay.excel.excelfiller.style.Color;

public class Font {
	private XSSFFont font;

	public Font(XSSFFont font) {
		this.font = font;
	}

	public Font boldweight(boolean boldFlag) {
		font.setBold(boldFlag);
		return this;
	}
	
	public Font italic(boolean italic) {
		font.setItalic(italic);
		return this; 
	}
	
	public Font fontHeightInPoints(short height) {
		font.setFontHeightInPoints(height);
		return this;
	}
	
	public Font color(Color color) {
		if (color.equals(Color.AUTOMATIC)) {
			font.setColor(XSSFFont.COLOR_NORMAL);
		} else {
			font.setColor(color.getIndex());
		}
		return this;
	}
	
	public Font underline(Underline underline) {
		font.setUnderline(underline.getLine());
		return this;
	}
}
