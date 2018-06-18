package com.bigcay.excel.excelfiller.template;

public class RowTemplate extends AbstractTemplate {

	private String arrayName;

	public RowTemplate(int row, int col, String templateType, String templateCode) {
		super(row, col, templateType, templateCode);
	}

	public String getArrayName() {
		return arrayName;
	}

	public void setArrayName(String arrayName) {
		this.arrayName = arrayName;
	}

}
