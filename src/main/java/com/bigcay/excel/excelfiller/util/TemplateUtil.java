package com.bigcay.excel.excelfiller.util;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;

import com.bigcay.excel.excelfiller.template.AbstractTemplate;
import com.bigcay.excel.excelfiller.template.CellTemplate;
import com.bigcay.excel.excelfiller.template.ColumnTemplate;
import com.bigcay.excel.excelfiller.template.MatrixTemplate;
import com.bigcay.excel.excelfiller.template.RowTemplate;

public class TemplateUtil {

	private final static Pattern TEMPLATE_ELEMENT_PATTERN = Pattern.compile("\\{\\<([^\\>]+?)\\>\\[(.+?)\\]([^\\}]*?)\\}");
	
	private final static String DOT_SEPARATOR = ".";
	private final static String COMMA_SEPARATOR = ",";
	private final static String COLON_SEPARATOR = ":";
	
	private final static String ROW_CHAR_NUM_CONFIG = "row_char_num";
	private final static String UNIT_ROW_HEIGHT_CONFIG = "unit_row_height";

	public final static String CELL_TYPE = "cell";

	public final static String COLUMN_LIST_TYPE = "column.list";

	public final static String ROW_ARRAY_TYPE = "row.array";

	public final static String MATRIX_DYADIC_ARRAY_TYPE = "matrix.2-dim-array";

	public static AbstractTemplate generateTemplate(int row, int col, String value, HSSFCellStyle cellStyle) {
		Matcher matcher = TEMPLATE_ELEMENT_PATTERN.matcher(value);
		if (matcher.find()) {
			if (matcher.groupCount() >= 2) {
				if (CELL_TYPE.equals(matcher.group(1))) {
					CellTemplate cellTemplate = new CellTemplate(row, col, matcher.group(1), matcher.group(2));
					
					String additionalConfig = matcher.group(3);
					if (additionalConfig.length() > 0) {
					    additionalConfig = additionalConfig.substring(1, additionalConfig.length() - 1);
                        String[] configureArray = additionalConfig.split(COMMA_SEPARATOR);
                        
                        boolean existRowCharNum = false;
                        boolean existUnitRowHeight = false;
                        
                        for (String config : configureArray) {
                            String[] configItemArray = config.split(COLON_SEPARATOR);
                            if (ROW_CHAR_NUM_CONFIG.equalsIgnoreCase(configItemArray[0])) {
                                existRowCharNum = true;
                                cellTemplate.setRowCharNumber(Integer.parseInt(configItemArray[1]));
                            } else if (UNIT_ROW_HEIGHT_CONFIG.equalsIgnoreCase(configItemArray[0])) {
                                existUnitRowHeight = true;
                                cellTemplate.setUnitRowHeightInPoint(Float.parseFloat(configItemArray[1]));
                            }
                        }
                        
                        if (existRowCharNum && existUnitRowHeight) {
                            cellTemplate.setAutoRowHeight(true);
                        }
					}
					
					return cellTemplate;
				} else if (COLUMN_LIST_TYPE.equals(matcher.group(1))) {
					int separatorIndex = matcher.group(2).indexOf(DOT_SEPARATOR);
					String listName = matcher.group(2).substring(0, separatorIndex);
					String methodName = matcher.group(2).substring(separatorIndex + 1);
					
					ColumnTemplate columnTemplate = new ColumnTemplate(row, col, matcher.group(1), matcher.group(2), cellStyle);
					columnTemplate.setListName(listName);
					columnTemplate.setMethodName(methodName);
					
					// Additional template configuration
                    String additionalConfig = matcher.group(3);
                    if (additionalConfig.length() > 0) {
                        additionalConfig = additionalConfig.substring(1, additionalConfig.length() - 1);
                        String[] configureArray = additionalConfig.split(COMMA_SEPARATOR);
                        
                        boolean existRowCharNum = false;
                        boolean existUnitRowHeight = false;
                        
                        for (String config : configureArray) {
                            String[] configItemArray = config.split(COLON_SEPARATOR);
                            if (ROW_CHAR_NUM_CONFIG.equalsIgnoreCase(configItemArray[0])) {
                                existRowCharNum = true;
                                columnTemplate.setRowCharNumber(Integer.parseInt(configItemArray[1]));
                            } else if (UNIT_ROW_HEIGHT_CONFIG.equalsIgnoreCase(configItemArray[0])) {
                                existUnitRowHeight = true;
                                columnTemplate.setUnitRowHeightInPoint(Float.parseFloat(configItemArray[1]));
                            }
                        }
                        
                        if (existRowCharNum && existUnitRowHeight) {
                            columnTemplate.setAutoRowHeight(true);
                        }
                    }
					
					return columnTemplate;
				} else if (ROW_ARRAY_TYPE.equals(matcher.group(1))) {
					RowTemplate rowTemplate = new RowTemplate(row, col, matcher.group(1), matcher.group(2));
					rowTemplate.setArrayName(matcher.group(2));
					return rowTemplate;
				} else if (MATRIX_DYADIC_ARRAY_TYPE.equals(matcher.group(1))) {
					MatrixTemplate matrixTemplate = new MatrixTemplate(row, col, matcher.group(1), matcher.group(2), cellStyle);
					return matrixTemplate;
				} else {
					return null;
				}
			} else {
				return null;
			}
		} else {
			return null;
		}
	}

	public static boolean isValidTemplateCode(String value) {
		Matcher matcher = TEMPLATE_ELEMENT_PATTERN.matcher(value);
		return matcher.find();
	}
	
    public static float getProperHeight4WrappedCell(String value, int rowCharNumber, float unitRowHeightInPoint) {
        int contentRowNumber = (String.valueOf(value).length() / rowCharNumber) + 1;
        return contentRowNumber * unitRowHeightInPoint;
    }
}
