package com.bigcay.excel.excelfiller.font;

import com.bigcay.excel.excelfiller.style.font.Font;

public class BoldFontUpdater implements IFontUpdater {

	@Override
	public void updateFont(Font font) {
		font.boldweight(true);
	}

}
