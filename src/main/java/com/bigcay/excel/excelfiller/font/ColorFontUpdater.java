package com.bigcay.excel.excelfiller.font;

import com.bigcay.excel.excelfiller.style.Color;
import com.bigcay.excel.excelfiller.style.font.Font;

public class ColorFontUpdater implements IFontUpdater {

	private Color color = Color.BLACK;
	
	@Override
	public void updateFont(Font font) {
		font.color(color);
	}

	public Color getColor() {
		return color;
	}

	public void setColor(Color color) {
		this.color = color;
	}

}
