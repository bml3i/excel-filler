package com.bigcay.excel.excelfiller.font;

import com.bigcay.excel.excelfiller.style.font.Font;
import com.bigcay.excel.excelfiller.style.font.Underline;

public class UnderlineFontUpdater implements IFontUpdater {

	private Underline underline = Underline.NONE;

	@Override
	public void updateFont(Font font) {
		font.underline(underline);
	}

	public Underline getUnderline() {
		return underline;
	}

	public void setUnderline(Underline underline) {
		this.underline = underline;
	}

}
