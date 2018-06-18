package com.bigcay.excel.excelfiller;

import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import com.bigcay.excel.excelfiller.beans.ReferenceCell;
import com.bigcay.excel.excelfiller.element.CellElement;
import com.bigcay.excel.excelfiller.element.ColumnElement;
import com.bigcay.excel.excelfiller.element.RowElement;
import com.bigcay.excel.excelfiller.element.SheetElement;
import com.bigcay.excel.excelfiller.style.DefaultStyle;
import com.bigcay.excel.excelfiller.template.AbstractTemplate;
import com.bigcay.excel.excelfiller.template.CellTemplate;
import com.bigcay.excel.excelfiller.template.ColumnTemplate;
import com.bigcay.excel.excelfiller.util.ExcelUtil;
import com.bigcay.excel.excelfiller.util.TemplateUtil;

public class Excel {
	private ExcelContext excelContext;

	public Excel() {
		this(new DefaultStyle());
	}

	public Excel(DefaultStyle defaultStyle) {
		this(null, null, defaultStyle);
	}

	public Excel(String filePath, Map<String, Object> excelData) {
		this(filePath, excelData, new DefaultStyle());
	}

	public Excel(String filePath, Map<String, Object> excelData, DefaultStyle defaultStyle) {
		XSSFWorkbook workbook;
		XSSFCellStyle tempCellStyle;
		XSSFFont tempFont;

		if (filePath == null || filePath.trim().length() == 0) {
			workbook = new XSSFWorkbook();
		} else {
			workbook = this.readExcelTemplate(filePath);
			if (workbook == null) {
				workbook = new XSSFWorkbook();
			}
		}

		excelContext = new ExcelContext(workbook);
		excelContext.setDefaultStyle(defaultStyle);

		if (excelData != null) {
			excelContext.setExcelData(excelData);
		}

		this.setWorkingSheet(0);

		tempCellStyle = workbook.createCellStyle();
		excelContext.setTempCellStyle(tempCellStyle);

		tempFont = workbook.createFont();
		excelContext.setTempFont(tempFont);

	}

	/**
	 * Read excel as template by filePath
	 * 
	 * @param filePath
	 * @return
	 */
	private XSSFWorkbook readExcelTemplate(String filePath) {
		XSSFWorkbook workbook = null;

		try {
			workbook = new XSSFWorkbook(new FileInputStream(filePath));
		} catch (Exception ex) {
			ex.printStackTrace();
		}

		return workbook;
	}
	
	public boolean saveExcel(String destFileName) {
		BufferedOutputStream os;
		try {
			os = new BufferedOutputStream(new FileOutputStream(destFileName));
			return saveExcel(os);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		return false;
	}
	
	public boolean saveExcel(OutputStream os) {
		boolean result = false;
		try {
			excelContext.getWorkbook().write(os);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				os.flush();
				os.close();
			} catch (Exception e) {
				result = false;
			}
		}
		return result;
	}

	public XSSFWorkbook exportWorkbook() {
		return excelContext.getWorkbook();
	}

	public SheetElement setWorkingSheet(int index) {
		excelContext.setWorkingSheet(ExcelUtil.getXSSFSheet(excelContext.getWorkbook(), index));
		return this.sheet(index);
	}

	public SheetElement sheet(int index) {
		SheetElement sheetElement = new SheetElement(ExcelUtil.getXSSFSheet(excelContext.getWorkbook(), index), excelContext);
		return sheetElement;
	}

	public CellElement cell(int row, int col) {
		return new CellElement(row, col, excelContext);
	}

	public ColumnElement column(int col) {
		ColumnElement columnElement = new ColumnElement(col, excelContext);
		return columnElement;
	}

	public ColumnElement column(int col, int startRow) {
		ColumnElement columnElement = new ColumnElement(col, startRow, excelContext);
		return columnElement;
	}

	public RowElement row(int row) {
		return new RowElement(row, excelContext);
	}

	public RowElement row(int row, int startCol) {
		return new RowElement(row, startCol, excelContext);
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void fillTemplateData() {
		for (int sheetIndex = 0; sheetIndex < excelContext.getWorkbook().getNumberOfSheets(); sheetIndex++) {
			XSSFSheet sheet = excelContext.getWorkbook().getSheetAt(sheetIndex);
			excelContext.setWorkingSheet(sheet);

			List<AbstractTemplate> templateList = new ArrayList<AbstractTemplate>();
			Map<String, List<ColumnTemplate>> columnListTemplateMap = new TreeMap<String, List<ColumnTemplate>>();

			for (Iterator rit = sheet.rowIterator(); rit.hasNext();) {
				XSSFRow row = (XSSFRow) rit.next();
				for (Iterator cit = row.cellIterator(); cit.hasNext();) {
					XSSFCell cell = (XSSFCell) cit.next();
					if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
						AbstractTemplate template = TemplateUtil.generateTemplate(cell.getRowIndex(), cell.getColumnIndex(),
								cell.getStringCellValue(), cell.getCellStyle());
						if (template != null) {
							templateList.add(template);
						}
					}
				}
			}
			
			int maxRowIndex = 0;
			// Deal with CellTemplate, RowTemplate, and generate ColumnTemplate list
			for (AbstractTemplate template : templateList) {
				if (TemplateUtil.CELL_TYPE.equals(template.getTemplateType())) {
				    CellTemplate cellTemplate = (CellTemplate) template;
					this.cell(template.getRow(), template.getCol()).setValue(excelContext.getExcelData().get(template.getTemplateCode()), cellTemplate);
				} else if (TemplateUtil.COLUMN_LIST_TYPE.equals(template.getTemplateType())) {
					ColumnTemplate columnListTemplate = (ColumnTemplate) template;
					List<ColumnTemplate> tmpColumnTemplateList = columnListTemplateMap.get(columnListTemplate.getListName());
					if (tmpColumnTemplateList == null) {
						tmpColumnTemplateList = new ArrayList<ColumnTemplate>();
						tmpColumnTemplateList.add(columnListTemplate);
						columnListTemplateMap.put(columnListTemplate.getListName(), tmpColumnTemplateList);
					} else {
						tmpColumnTemplateList.add(columnListTemplate);
					}
				} else if (TemplateUtil.ROW_ARRAY_TYPE.equals(template.getTemplateType())) {
					this.row(template.getRow(), template.getCol()).setValue(
							(Object[]) excelContext.getExcelData().get(template.getTemplateCode()), true);
				} else if(TemplateUtil.MATRIX_DYADIC_ARRAY_TYPE.equals(template.getTemplateType())) {
					Object[] dyadicArray = (Object[]) excelContext.getExcelData().get(template.getTemplateCode());
					
					ReferenceCell refCell = new ReferenceCell(template.getRow(), template.getCol(), template.getTemplateStyle());
					
					if (template.getRow() > maxRowIndex) {
						maxRowIndex = template.getRow();
					}
					for(int i = 0; i < dyadicArray.length; i ++) {
						Object[] rowData = (Object[])dyadicArray[i];
						this.row(maxRowIndex + i, template.getCol()).setValue(rowData, refCell);
					}
					maxRowIndex += dyadicArray.length;
				}
			}

			// Deal with ColumnTemplate
			maxRowIndex = 0;
			for (String listName : columnListTemplateMap.keySet()) {
				List<Object> listObjectList = (ArrayList<Object>) excelContext.getExcelData().get(listName);

				if (listObjectList == null || listObjectList.size() == 0) {
					continue;
				}

				List<ColumnTemplate> columnTemplateList = columnListTemplateMap.get(listName);
				Map<String, Object[]> columnDataMap = new HashMap<String, Object[]>();

				if (listObjectList != null) {
					for (int index = 0; index < listObjectList.size(); index++) {
						Object dataObj = listObjectList.get(index);
						Class clazz = dataObj.getClass();

						for (ColumnTemplate columnTemplate : columnTemplateList) {
							String methodName = columnTemplate.getMethodName();

							Object objValue = new Object();
							try {
								Method method = clazz.getDeclaredMethod(methodName);
								objValue = method.invoke(dataObj);
							} catch (Exception ex) {
							}

							if (index == 0) {
								Object[] dataObjs = new Object[listObjectList.size()];
								dataObjs[index] = objValue;
								columnDataMap.put(columnTemplate.getTemplateCode(), dataObjs);
							} else {
								Object[] dataObjs = columnDataMap.get(columnTemplate.getTemplateCode());
								dataObjs[index] = objValue;
							}
						}
					}
				}

				int maxRowIndexTmp = 0;
				int formerRowIndex = 0;
				for (ColumnTemplate columnTemplate : columnTemplateList) {
					Object[] columnDataArray = columnDataMap.get(columnTemplate.getTemplateCode());
					if (columnDataArray != null) {
						if (formerRowIndex != columnTemplate.getRow() && maxRowIndexTmp > maxRowIndex) {
							maxRowIndex = maxRowIndexTmp;
						}
						if (columnTemplate.getRow() > maxRowIndex) {
							maxRowIndex = columnTemplate.getRow();
						}
						this.column(columnTemplate.getCol(), maxRowIndex)
								.setValue(columnDataArray, true, columnTemplate);
						formerRowIndex = columnTemplate.getRow();
						if ((maxRowIndex + columnDataArray.length) > maxRowIndexTmp) {
							maxRowIndexTmp = maxRowIndex + columnDataArray.length;
						}
					}
				}
				maxRowIndex = maxRowIndexTmp;
			}

			// Clear all the template expression
			for (AbstractTemplate template : templateList) {
				CellElement cellElem = this.cell(template.getRow(), template.getCol());
				XSSFCell cell = cellElem.getWorkingCells().get(0);

				if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
					if (TemplateUtil.isValidTemplateCode(cell.getStringCellValue())) {
						cellElem.setValue(null);
					}
				}
			}
		}
	}

}
