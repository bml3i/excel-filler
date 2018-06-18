package com.bigcay.excel.excelfiller.font;

import com.bigcay.excel.excelfiller.style.font.Font;

public class HeightFontUpdater implements IFontUpdater {

	private int height = 12; 
	
	@Override
	public void updateFont(Font font) {
		font.fontHeightInPoints((short) height);
	}

	public int getHeight() {
		return height;
	}

	public void setHeight(int height) {
		this.height = height;
	}

}
