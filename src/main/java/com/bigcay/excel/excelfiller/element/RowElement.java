package com.bigcay.excel.excelfiller.element;

import org.apache.poi.hssf.usermodel.HSSFRow;

import com.bigcay.excel.excelfiller.ExcelContext;
import com.bigcay.excel.excelfiller.beans.ReferenceCell;

public class RowElement extends AbstractElement {

	private HSSFRow row;

	private int startCol;

	public RowElement(ExcelContext excelContext) {
		super(excelContext);
	}

	public RowElement(int row, ExcelContext excelContext) {
		this(row, 0, excelContext);
	}

	public RowElement(int row, int startCol, ExcelContext excelContext) {
		super(excelContext);
		this.row = this.getRow(row);
		this.startCol = startCol;
	}

	public RowElement setValue(Object[] rowData) {
		setValue(rowData, startCol);
		return this;
	}

	public RowElement setValue(Object[] rowData, boolean applyCellStyleFlag) {
		setValue(rowData, startCol, applyCellStyleFlag);
		return this;
	}
	
	public RowElement setValue(Object[] rowData, ReferenceCell refCell) {
	    insertData(rowData, row, startCol, refCell);
	    return this; 
	}

	public RowElement setValue(Object[] rowData, int startCol) {
		insertData(rowData, row, startCol, false);
		return this;
	}

	public RowElement setValue(Object[] rowData, int startCol, boolean applyCellStyleFlag) {
		insertData(rowData, row, startCol, applyCellStyleFlag);
		return this;
	}

	private void insertData(Object[] rowData, HSSFRow row, int startCol, boolean applyCellStyleFlag) {
		for (int index = 0; index < rowData.length; index++) {
			CellElement cellElement = new CellElement(row.getRowNum(), startCol + index, excelContext);
			cellElement.setValue(rowData[index]);

            // Copy and apply CellStyle from the genetic cell if applyCellStyleFlag is true
			if (applyCellStyleFlag) {
				cellElement.applyCellStyle(row.getRowNum(), startCol);
			}
		}
	}
	
    private void insertData(Object[] rowData, HSSFRow row, int startCol, ReferenceCell refCell) {
        for (int index = 0; index < rowData.length; index++) {
            CellElement cellElement = new CellElement(row.getRowNum(), startCol + index, excelContext);
            cellElement.setValue(rowData[index]);
            
            // Copy and apply CellStyle from the reference cell
            cellElement.applyCellStyle(refCell.getCellStyle());
        }
    }
}
