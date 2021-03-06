package com.bigcay.excel.excelfiller.element;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;

import com.bigcay.excel.excelfiller.ExcelContext;
import com.bigcay.excel.excelfiller.font.BoldFontUpdater;
import com.bigcay.excel.excelfiller.font.ColorFontUpdater;
import com.bigcay.excel.excelfiller.font.HeightFontUpdater;
import com.bigcay.excel.excelfiller.font.IFontUpdater;
import com.bigcay.excel.excelfiller.font.ItalicFontUpdater;
import com.bigcay.excel.excelfiller.font.UnderlineFontUpdater;
import com.bigcay.excel.excelfiller.style.Align;
import com.bigcay.excel.excelfiller.style.Color;
import com.bigcay.excel.excelfiller.style.FillPattern;
import com.bigcay.excel.excelfiller.style.font.Font;
import com.bigcay.excel.excelfiller.style.font.Underline;
import com.bigcay.excel.excelfiller.template.AbstractTemplate;
import com.bigcay.excel.excelfiller.util.ExcelUtil;
import com.bigcay.excel.excelfiller.util.TemplateUtil;

public class CellElement extends AbstractElement {

	private List<XSSFCell> workingCells = new ArrayList<XSSFCell>();
	
	private static BoldFontUpdater boldFont = new BoldFontUpdater();
	private static ItalicFontUpdater italicFont = new ItalicFontUpdater();
	private static HeightFontUpdater heightFont = new HeightFontUpdater(); 
	private static ColorFontUpdater colorFont = new ColorFontUpdater();
	private static UnderlineFontUpdater underlineFont = new UnderlineFontUpdater();

	public CellElement(ExcelContext excelContext) {
		super(excelContext);
	}

	public List<XSSFCell> getWorkingCells() {
		return workingCells;
	}

	public CellElement(int row, int col, ExcelContext excelContext) {
		super(excelContext);
		this.add(row, col);
	}

	protected CellElement add(int row, int col) {
		XSSFCell cell = getCell(row, col);
		workingCells.add(cell);
		return this;
	}

	protected XSSFCell getCell(int row, int col) {
		return ExcelUtil.getXSSFCell(this.workingSheet, row, col);
	}

	public CellElement setValue(Object value) {
		for (XSSFCell cell : workingCells) {
			this.setCellValue(cell, value, null);
		}
		return this;
	}
	
    public CellElement setValue(Object value, AbstractTemplate template) {
        for (XSSFCell cell : workingCells) {
            this.setCellValue(cell, value, null);
        }
        
        if (template.isAutoRowHeight()) {
            float properHeightInPoint = TemplateUtil.getProperHeight4WrappedCell(String.valueOf(value),
                    template.getRowCharNumber(),
                    template.getUnitRowHeightInPoint());
            this.setHigherHeight(properHeightInPoint);
        }
        
        return this;
    }

	private void setCellValue(XSSFCell cell, Object value, String pattern) {
		if (value instanceof Double || value instanceof Float
				|| value instanceof Long || value instanceof Integer
				|| value instanceof Short || value instanceof BigDecimal
				|| value instanceof Byte) {
			cell.setCellValue(convert2Double(value));
			cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
			cell.setCellType(XSSFCell.CELL_TYPE_BOOLEAN);
		} else {
			if (value instanceof Date) {
				if (pattern == null || pattern.trim().length() == 0) {
					pattern = excelContext.getDefaultStyle().getDatePattern();
				}
				cell.setCellValue((Date) value);
			} else {
				cell.setCellValue(new XSSFRichTextString(value == null ? ""
						: value.toString()));
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			}
		}

		if (pattern != null) {
			this.dataFormat(pattern);
		}
	}

	// TO-DO
	public CellElement dataFormat(String format) {
		short index = excelContext.getWorkbook().createDataFormat().getFormat(format);
		for (XSSFCell cell : workingCells) {
			XSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			if (index == -1) {
				XSSFDataFormat dataFormat = excelContext.getWorkbook()
						.createDataFormat();
				index = dataFormat.getFormat(format);
			}
			tempCellStyle.setDataFormat(index);
			updateCellStyle(cell);
		}
		return this;
	}

	public CellElement applyCellStyle(int row, int col) {
		XSSFCellStyle geneticCellStyle = this.getCell(row, col).getCellStyle();
		for (XSSFCell cell : workingCells) {
			cell.setCellStyle(geneticCellStyle);
		}
		return this;
	}
	
	public CellElement applyCellStyle(XSSFCellStyle cellStyle) {
		for (XSSFCell cell : workingCells) {
			cell.setCellStyle(cellStyle);
		}
		return this;
	}

	public CellElement align(Align align) {
		for (XSSFCell cell : workingCells) {
			XSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setAlignment(align.getAlignment());
			updateCellStyle(cell);
		}
		return this;
	}

	public CellElement bgColor(Color bg) {
		return bgColor(bg, FillPattern.SOLID_FOREGROUND);
	}

	public CellElement bgColor(Color bg, FillPattern fillPattern) {
		for (XSSFCell cell : workingCells) {
			XSSFCellStyle style = cell.getCellStyle();
			tempCellStyle.cloneStyleFrom(style);
			tempCellStyle.setFillPattern(fillPattern.getFillPattern());
			tempCellStyle.setFillForegroundColor(bg.getIndex());
			updateCellStyle(cell);
		}
		return this;
	}
	
	public CellElement height(float height){
		for (XSSFCell cell : workingCells) {
			XSSFRow row = getRow(cell.getRowIndex());
			row.setHeightInPoints(height);
		}
		return this;
	}
	
    public CellElement setHigherHeight(float height) {
        for (XSSFCell cell : workingCells) {
            XSSFRow row = getRow(cell.getRowIndex());
            if(row.getHeightInPoints() < height) {
                row.setHeightInPoints(height);                
            }
        }
        return this;
    }
	
	public CellElement bold() {
		font(boldFont);
		return this; 
	}
	
	public CellElement italic() {
		font(italicFont);
		return this; 
	}
	
	public CellElement color(Color color) {
		colorFont.setColor(color);
		font(colorFont);
		return this;
	}
	
	public CellElement fontHeightInPoint(int height) {
		heightFont.setHeight(height);
		font(heightFont);
		return this;
	}
	
	public CellElement underline(Underline underline) {
		underlineFont.setUnderline(underline);
		font(underlineFont);
		return this;
	}
	
	public CellElement font(IFontUpdater fontUpdater) {
		Map<Integer, XSSFFont> fontPool = excelContext.getFontPool();
		for (XSSFCell cell : workingCells) {
			XSSFFont font = cell.getCellStyle().getFont();
			copyFont(font, tempFont);
			fontUpdater.updateFont(new Font(tempFont));
			int fontHash = tempFont.hashCode() - tempFont.getIndex();
			tempCellStyle.cloneStyleFrom(cell.getCellStyle());
			if (fontPool.containsKey(fontHash)) {
				tempCellStyle.setFont(fontPool.get(fontHash));
			} else {
				XSSFFont newFont = workbook.createFont();
				copyFont(tempFont, newFont);
				tempCellStyle.setFont(newFont);
				int newFontHash = newFont.hashCode() - newFont.getIndex();
				fontPool.put(newFontHash, newFont);
			}
			updateCellStyle(cell);
		}
		return this;
	}
	
	private void copyFont(XSSFFont sourceFont, XSSFFont targetFont) {
		targetFont.setBold(sourceFont.getBold());
		targetFont.setCharSet(sourceFont.getCharSet());
		targetFont.setColor(sourceFont.getColor());
		targetFont.setFontHeight(sourceFont.getFontHeight());
		targetFont.setFontHeightInPoints(sourceFont.getFontHeightInPoints());
		targetFont.setFontName(sourceFont.getFontName());
		targetFont.setItalic(sourceFont.getItalic());
		targetFont.setStrikeout(sourceFont.getStrikeout());
		targetFont.setTypeOffset(sourceFont.getTypeOffset());
		targetFont.setUnderline(sourceFont.getUnderline());
	}

	private void updateCellStyle(XSSFCell cell) {
		Map<Integer, XSSFCellStyle> stylePool = excelContext.getStylePool();
		int tempStyleHash = tempCellStyle.hashCode() - tempCellStyle.getIndex();
		if (stylePool.containsKey(tempStyleHash)) {
			cell.setCellStyle(stylePool.get(tempStyleHash));
		} else {
			XSSFCellStyle newStyle = workbook.createCellStyle();
			newStyle.cloneStyleFrom(tempCellStyle);
			cell.setCellStyle(newStyle);
			int newStyleHash = newStyle.hashCode() - newStyle.getIndex();
			stylePool.put(newStyleHash, newStyle);
		}
	}

	private double convert2Double(Object obj) {
		double result = 0;
		if (obj != null) {
			try {
				result = Double.parseDouble(obj.toString());
			} catch (Exception ex) {
			}
		}
		return result;
	}

}
