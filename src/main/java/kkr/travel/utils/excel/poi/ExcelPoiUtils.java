package kkr.travel.utils.excel.poi;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.TimeZone;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import kkr.common.errors.BaseException;
import kkr.common.errors.TechnicalException;
import kkr.common.utils.excel.ExcelPosition;

public class ExcelPoiUtils {
	private static final long MS_1900_1970 = 2209078800000L;

	public static final DateFormat DATE_FORMAT_DATE = new SimpleDateFormat();

	public static boolean isEmptyRow(PoiWorkbook poiWorkbook, Row row) throws BaseException {
		if (row == null) {
			return true;
		}

		int nColumns = row.getLastCellNum();
		for (int iCol = row.getFirstCellNum(); iCol < nColumns; iCol++) {
			Cell cell = row.getCell(iCol);
			if (!isEmptyCell(poiWorkbook, cell)) {
				return false;
			}
		}
		return true;
	}

	public static boolean isEmptyCell(PoiWorkbook poiWorkbook, Cell cell) throws BaseException {
		if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
			return true;
		}

		CellValue cellValue = null;
		try {
			cellValue = poiWorkbook.getEvaluator().evaluate(cell);
		} catch (Exception ex) {
			return false;
		}

		if (cellValue == null) {
			return true;
		}

		String strValue = cellValue.formatAsString().trim();
		if (!strValue.isEmpty() && strValue.startsWith("\"") && strValue.endsWith("\"")) {
			strValue = strValue.length() > 2 ? strValue.substring(1, strValue.length() - 1).trim() : "";
		}
		return strValue.isEmpty();
	}

	public static Object readValue(ExcelPosition excelPosition, PoiWorkbook poiWorkbook, Cell cell) throws BaseException {

		if (cell == null) {
			return null;
		}

		CellValue cellValue = null;
		try {
			cellValue = poiWorkbook.getEvaluator().evaluate(cell);
		} catch (Exception ex) {
			String message = "Cannot evaluate the cell value: " + cell.toString();
			if (poiWorkbook.getEvaluator() instanceof FormulaEvaluator) {
				message += " Check for empty arguments in the expression.";
			}

			throw new TechnicalException(message, ex);
		}

		if (cellValue == null) {
			return null;
		}

		switch (cellValue.getCellType()) {
			case Cell.CELL_TYPE_STRING :
				return emptyToNull(cellValue.getStringValue());
			case Cell.CELL_TYPE_NUMERIC : {
				Date date = toDate(cell, cellValue);
				if (date != null) {
					return date;
				}
				double dblValue = cellValue.getNumberValue();
				int intValue = (int) dblValue;
				double dblIntValue = (double) intValue;
				if (dblIntValue == dblValue) {
					return new Integer(intValue);
				} else {
					return new Double(dblValue);
				}
			}
			case Cell.CELL_TYPE_BOOLEAN :
				return new Boolean(cellValue.getBooleanValue());
			case Cell.CELL_TYPE_BLANK :
				return null;
			default :
				return emptyToNull(cellValue.formatAsString());
		}
	}

	private static String emptyToNull(String value) {
		return value.trim().isEmpty() ? null : value;
	}

	public static void writeValue(PoiWorkbook poiWorkbook, Cell cell, Object value) {
		if (value != null) {
			if (value instanceof Number) {
				cell.setCellValue(((Number) value).doubleValue());
			} else if (value instanceof Date) {
				CellStyle cellStyle = poiWorkbook.getWorkbook().createCellStyle();
				DataFormat dataFormat = poiWorkbook.getWorkbook().createDataFormat();
				short idFormat = dataFormat.getFormat("dd/mm/yyyy");
				cellStyle.setDataFormat(idFormat);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(((Date) value));
			} else if (value instanceof Boolean) {
				cell.setCellValue(((Boolean) value));
			} else {
				cell.setCellValue(String.valueOf(value));
			}
		} else {
			cell.setCellValue((String) null);
		}
	}

	public static String readValueString(ExcelPosition excelPosition, PoiWorkbook poiWorkbook, Cell cell) throws BaseException {
		if (cell == null) {
			return null;
		}

		Object object = readValue(excelPosition, poiWorkbook, cell);
		if (object == null) {
			return null;
		}

		if (object instanceof Date) {
			return DATE_FORMAT_DATE.format(object);
		} else if (object instanceof Integer) {
			return String.format("%d", object);
		} else if (object instanceof Double) {
			return String.format("%f", object);
		} else if (object instanceof String) {
			return (String) object;
		} else {
			return object.toString();
		}
	}

	private static Date toDate(Cell cell, CellValue cellValue) {
		if (cellValue == null) {
			return null;
		}
		CellStyle cellStyle = cell.getCellStyle();
		String dataFormatString = cellStyle.getDataFormatString();

		if (dataFormatString == null || dataFormatString.isEmpty()) {
			return null;
		}

		if (dataFormatString.matches(".*[dMyhmsS].*")) {
			double value = cellValue.getNumberValue();
			long ms = (long) (value * 24 * 60 * 60 * 1000);
			long msDelta = ms - MS_1900_1970;
			Date date = new Date(msDelta);
			Calendar calendar = new GregorianCalendar();
			calendar.setTimeZone(TimeZone.getTimeZone("GMT"));
			calendar.setTimeInMillis(msDelta);
			date = calendar.getTime();
			return date;
		} else {
			return null;
		}
	}

	public static String getValueString(ExcelPosition excelPosition, PoiWorkbook poiWorkbook, Row row, int column) {
		Cell cell = row.getCell(column);

		if (cell == null) {
			return null;
		}
		CellValue cellValue = poiWorkbook.getEvaluator().evaluate(cell);
		if (cellValue == null) {
			return null;
		}
		if (cellValue.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			double doubleValue = cellValue.getNumberValue();
			int intValue = (int) doubleValue;
			if (doubleValue == ((double) intValue)) {
				return String.valueOf(intValue);
			} else {
				return String.valueOf(doubleValue);
			}
		} else if (cellValue.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
			boolean booleanValue = cellValue.getBooleanValue();
			return String.valueOf(booleanValue);
		} else {
			return cellValue.getStringValue();
		}
	}

	public static void createSheetColors(PoiWorkbook poiWorkbook) {
		Sheet sheet = poiWorkbook.getWorkbook().createSheet("COLORS");

		Row row;
		Cell cell;
		CellStyle cellStyle;

		int iRow = 0;
		for (IndexedColors indexedColor : IndexedColors.values()) {
			row = sheet.getRow(iRow);
			if (row == null) {
				row = sheet.createRow(iRow);
			}

			cell = row.getCell(0, Row.CREATE_NULL_AS_BLANK);
			cell.setCellValue("");
			cellStyle = poiWorkbook.getWorkbook().createCellStyle();
			cellStyle.setFillForegroundColor(indexedColor.getIndex());
			cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			cell.setCellStyle(cellStyle);

			cell = row.getCell(1, Row.CREATE_NULL_AS_BLANK);
			cell.setCellValue(indexedColor.name());

			iRow++;
		}
	}
}
