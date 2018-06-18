package com.bigcay.excel.excelfiller.font;

import com.bigcay.excel.excelfiller.style.font.Font;

public class ItalicFontUpdater implements IFontUpdater {

	@Override
	public void updateFont(Font font) {
		font.italic(true);
	}

}
