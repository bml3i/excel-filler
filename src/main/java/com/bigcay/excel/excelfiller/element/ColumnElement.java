package com.bigcay.excel.excelfiller.element;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import com.bigcay.excel.excelfiller.ExcelContext;
import com.bigcay.excel.excelfiller.template.AbstractTemplate;
import com.bigcay.excel.excelfiller.util.TemplateUtil;

public class ColumnElement extends AbstractElement {

	private int col = 0;

	public int getCol() {
		return col;
	}

	private int startRow = 0;

	public ColumnElement(ExcelContext excelContext) {
		super(excelContext);
	}

	public ColumnElement(int col, int startRow, ExcelContext excelContext) {
		super(excelContext);
		this.col = col;
		this.startRow = startRow;
	}

	public ColumnElement(int col, ExcelContext excelContext) {
		this(col, 0, excelContext);
	}

	public ColumnElement setValue(Object[] columnData) {
		setValue(columnData, this.startRow, false);
		return this;
	}

	public ColumnElement setValue(Object[] columnData, boolean applyCellStyleFlag) {
		setValue(columnData, this.startRow, applyCellStyleFlag);
		return this;
	}
	
    public ColumnElement setValue(Object[] columnData, boolean applyCellStyleFlag, AbstractTemplate template) {
        if (template.isAutoRowHeight()) {
            insertData(columnData, this.col, startRow, applyCellStyleFlag, template.getRowCharNumber(), template.getUnitRowHeightInPoint(), template.getTemplateStyle());
        } else {
            insertData(columnData, this.col, startRow, applyCellStyleFlag, template.getTemplateStyle());
        }
        
        return this;
    }

	public ColumnElement setValue(Object[] columnData, int startRow) {
		insertData(columnData, this.col, startRow, false);
		return this;
	}

	public ColumnElement setValue(Object[] columnData, int startRow, boolean applyCellStyleFlag) {
		insertData(columnData, this.col, startRow, applyCellStyleFlag);
		return this;
	}

	private void insertData(Object[] columnData, int col, int startRow, boolean applyCellStyleFlag) {
		for (int index = 0; index < columnData.length; index++) {
			CellElement cellElement = new CellElement(this.startRow + index, this.col, excelContext);
			cellElement.setValue(columnData[index]);

			// Copy and apply CellStyle from the genetic cell if
			// applyCellStyleFlag is true
			if (applyCellStyleFlag) {
				cellElement.applyCellStyle(startRow, col);
			}
		}
	}
	
	private void insertData(Object[] columnData, int col, int startRow, boolean applyCellStyleFlag, XSSFCellStyle cellStyle) {
		for (int index = 0; index < columnData.length; index++) {
			CellElement cellElement = new CellElement(this.startRow + index, this.col, excelContext);
			cellElement.setValue(columnData[index]);
			
			// Copy and apply CellStyle from the genetic cell if
			// applyCellStyleFlag is true
			if (applyCellStyleFlag) {
				cellElement.applyCellStyle(cellStyle);
			}
		}
	}
	
    private void insertData(Object[] columnData, int col, int startRow, boolean applyCellStyleFlag, int rowCharNumber,
            float unitRowHeightInPoint, XSSFCellStyle cellStyle) {
        for (int index = 0; index < columnData.length; index++) {
            CellElement cellElement = new CellElement(this.startRow + index, this.col, excelContext);
            Object value = columnData[index];
            cellElement.setValue(value);
            
            // Copy and apply CellStyle from the genetic cell if
            // applyCellStyleFlag is true
            if (applyCellStyleFlag) {
                cellElement.applyCellStyle(cellStyle);
            }
            
            float properHeightInPoint = TemplateUtil.getProperHeight4WrappedCell(String.valueOf(value), rowCharNumber, unitRowHeightInPoint);
            cellElement.setHigherHeight(properHeightInPoint);
        }
    }

}
