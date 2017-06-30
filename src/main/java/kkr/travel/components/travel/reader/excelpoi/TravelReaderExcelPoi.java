package kkr.travel.components.travel.reader.excelpoi;

import java.net.MalformedURLException;
import java.net.URL;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import kkr.common.errors.BaseException;
import kkr.common.errors.ExcelException;
import kkr.common.utils.excel.ExcelPosition;
import kkr.travel.components.travel.data.Constants;
import kkr.travel.components.travel.data.LinkData;
import kkr.travel.components.travel.data.TravelData;
import kkr.travel.components.travel.reader.TravelReader;
import kkr.travel.utils.excel.poi.ExcelPoiUtils;
import kkr.travel.utils.excel.poi.PoiManipulator;
import kkr.travel.utils.excel.poi.PoiWorkbook;

public class TravelReaderExcelPoi extends TravelReaderExcelPoiFwk implements TravelReader, Constants {
	private static final Logger LOG = Logger.getLogger(TravelReaderExcelPoi.class);

	private static final DateFormat DATE_FORMAT = new SimpleDateFormat("yyyyMMdd");

	public Collection<TravelData> readTravelData() throws BaseException {
		LOG.trace("BEGIN");
		try {
			testConfigured();

			Collection<TravelData> data = new ArrayList<TravelData>();

			PoiWorkbook poiWorkbook = null;

			try {
				ExcelPosition excelPosition = new ExcelPosition();
				poiWorkbook = PoiManipulator.openWorkbook(file);
				excelPosition.setFile(file);

				Sheet sheet = poiWorkbook.getWorkbook().getSheet(this.sheet);
				if (sheet == null) {
					throw new ExcelException(excelPosition, "The sheet not found in the workbook: " + this.sheet);
				}
				excelPosition.setSheet(this.sheet);

				Map<String, Integer> header = readHeader(excelPosition, poiWorkbook, sheet);

				data = readData(excelPosition, poiWorkbook, sheet, header);

				PoiManipulator.closeWorkbook(poiWorkbook);
				poiWorkbook = null;
			} finally {
				try {
					PoiManipulator.closeWorkbook(poiWorkbook);
				} catch (Exception ex) {
					// nothing to do
				}
			}

			LOG.trace("OK");
			return data;
		} finally {
			LOG.trace("END");
		}
	}

	private Map<String, Integer> readHeader(ExcelPosition excelPositionSheet, PoiWorkbook poiWorkbook, Sheet sheet) throws BaseException {
		LOG.trace("BEGIN");
		try {

			Map<String, Integer> retval = new HashMap<String, Integer>();

			ExcelPosition excelPosition = excelPositionSheet.clone();

			Row row = sheet.getRow(0);
			excelPosition.setRow(0);

			if (row == null) {
				throw new ExcelException(excelPosition, "The sheet not found in the workbook: " + this.sheet);
			}

			int nColumns = row.getLastCellNum();
			for (int i = 0; i <= nColumns; i++) {
				excelPosition.setColumn(i);
				Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
				String value = ExcelPoiUtils.readValueString(excelPosition, poiWorkbook, cell);
				if (value == null) {
					value = "";
				} else {
					value = value.trim();
				}

				if (isEmpty(value)) {
					continue;
				}

				boolean found = false;

				found |= readHeaderColumn(excelPosition, retval, COLUMN_NAME, i, value);
				found |= readHeaderColumn(excelPosition, retval, COLUMN_DATE_FROM, i, value);
				found |= readHeaderColumn(excelPosition, retval, COLUMN_DATE_TO, i, value);
				found |= readHeaderColumn(excelPosition, retval, COLUMN_IMMAGE_MAP, i, value);
				found |= readHeaderColumn(excelPosition, retval, COLUMN_IMMAGE_TITLE, i, value);
				found |= readHeaderColumn(excelPosition, retval, COLUMN_LINK_JOURNAL, i, value);
				found |= readHeaderColumn(excelPosition, retval, COLUMN_LINK_PHOTOS, i, value);
				found |= readHeaderColumn(excelPosition, retval, COLUMN_LINK_TRACES, i, value);
				found |= readHeaderColumn(excelPosition, retval, COLUMN_PEOPLE, i, value);
				found |= readHeaderColumn(excelPosition, retval, COLUMN_PLACES, i, value);

				if (!found) {
					throw new ExcelException(excelPosition, "Unknown column: " + value);
				}
			}

			if (!retval.containsKey(COLUMN_NAME)) {
				throw new ExcelException(excelPosition, "Missing mandatory column: " + COLUMN_NAME);
			}
			if (!retval.containsKey(COLUMN_DATE_FROM)) {
				throw new ExcelException(excelPosition, "Missing mandatory column: " + COLUMN_DATE_FROM);
			}

			LOG.trace("OK");
			return retval;
		} finally {
			LOG.trace("END");
		}
	}

	private boolean readHeaderColumn(ExcelPosition excelPosition, Map<String, Integer> map, String columnName, int columnIndex, String value)
			throws BaseException {
		if (columnName.equals(value)) {
			if (map.containsKey(columnName)) {
				throw new ExcelException(excelPosition, "The column already exists in the file: " + columnName);
			}
			map.put(columnName, columnIndex);
			return true;
		}
		return false;
	}

	private Collection<TravelData> readData(ExcelPosition excelPositionSheet, PoiWorkbook poiWorkbook, Sheet sheet, Map<String, Integer> header)
			throws BaseException {
		LOG.trace("BEGIN");
		try {
			Collection<TravelData> retval = new ArrayList<TravelData>();

			ExcelPosition excelPosition = excelPositionSheet.clone();

			int nrows = sheet.getLastRowNum();

			for (int i = 1; i <= nrows; i++) {
				Row row = sheet.getRow(i);
				excelPosition.setRow(i);

				if (ExcelPoiUtils.isEmptyRow(poiWorkbook, row)) {
					continue;
				}

				TravelData travelData = readLine(excelPosition, poiWorkbook, row, header);

				retval.add(travelData);
			}

			LOG.trace("OK");
			return retval;
		} finally {
			LOG.trace("END");
		}

	}

	private TravelData readLine(ExcelPosition excelPositionRow, PoiWorkbook poiWorkbook, Row row, Map<String, Integer> header) throws BaseException {
		ExcelPosition excelPosition = excelPositionRow.clone();

		Cell cell;
		Integer index;

		//
		// NAME
		//
		index = header.get(COLUMN_NAME);
		excelPosition.setColumn(index);
		cell = row.getCell(index);

		String valueName = ExcelPoiUtils.readValueString(excelPosition, poiWorkbook, cell);
		if (isEmpty(valueName)) {
			throw new ExcelException(excelPosition, COLUMN_NAME + " is mandatory");
		}

		TravelData retval = new TravelData(valueName);

		//
		// DATE FROM
		//
		index = header.get(COLUMN_DATE_FROM);
		excelPosition.setColumn(index);
		cell = row.getCell(index);

		Object valueDateFrom = ExcelPoiUtils.readValue(excelPosition, poiWorkbook, cell);
		if (isEmpty(valueDateFrom)) {
			throw new ExcelException(excelPosition, COLUMN_DATE_FROM + " is mandatory");
		}

		Date date = toDate(valueDateFrom);
		if (date == null) {
			throw new ExcelException(excelPosition, COLUMN_DATE_FROM + " value is not a date: " + String.valueOf(valueDateFrom));
		}
		retval.setDateFrom(date);

		//
		// DATE TO
		//
		index = header.get(COLUMN_DATE_TO);
		if (index != null) {
			excelPosition.setColumn(index);
			cell = row.getCell(index);

			Object valueDateTo = ExcelPoiUtils.readValue(excelPosition, poiWorkbook, cell);
			if (!isEmpty(valueDateTo)) {
				date = toDate(valueDateTo);
				if (date == null) {
					throw new ExcelException(excelPosition, COLUMN_DATE_TO + " value is not a date: " + String.valueOf(valueDateTo));
				}
				retval.setDateTo(date);
			}
		}

		//
		// PEOPLE
		//
		index = header.get(COLUMN_PEOPLE);
		if (index != null) {
			excelPosition.setColumn(index);
			cell = row.getCell(index);
			String value = ExcelPoiUtils.readValueString(excelPosition, poiWorkbook, cell);
			Collection<String> list = toListString(excelPosition, COLUMN_PEOPLE, value);
			retval.getPeople().addAll(list);
		}

		//
		// PLACES
		//
		index = header.get(COLUMN_PLACES);
		if (index != null) {
			excelPosition.setColumn(index);
			cell = row.getCell(index);
			String value = ExcelPoiUtils.readValueString(excelPosition, poiWorkbook, cell);
			Collection<String> list = toListString(excelPosition, COLUMN_PLACES, value);
			retval.getPlaces().addAll(list);
		}

		//
		// IMMAGE_TITLE
		//
		index = header.get(COLUMN_IMMAGE_TITLE);
		if (index != null) {
			excelPosition.setColumn(index);
			cell = row.getCell(index);
			String value = ExcelPoiUtils.readValueString(excelPosition, poiWorkbook, cell);
			if (!isEmpty(value)) {
				try {
					URL url = new URL(value);
					retval.setImmageTitle(url);
				} catch (MalformedURLException ex) {
					throw new ExcelException(excelPosition, COLUMN_IMMAGE_TITLE + " value is not a URL: " + value);
				}
			}
		}

		//
		// IMMAGE_MAP
		//
		index = header.get(COLUMN_IMMAGE_MAP);
		if (index != null) {
			excelPosition.setColumn(index);
			cell = row.getCell(index);
			String value = ExcelPoiUtils.readValueString(excelPosition, poiWorkbook, cell);
			if (!isEmpty(value)) {
				try {
					URL url = new URL(value);
					retval.setImmageMap(url);
				} catch (MalformedURLException ex) {
					throw new ExcelException(excelPosition, COLUMN_IMMAGE_MAP + " value is not a URL: " + value);
				}
			}
		}

		//
		// LINK_PHOTOS
		//
		index = header.get(COLUMN_LINK_PHOTOS);
		if (index != null) {
			excelPosition.setColumn(index);
			cell = row.getCell(index);
			String value = ExcelPoiUtils.readValueString(excelPosition, poiWorkbook, cell);
			Collection<String> linksString = toListString(excelPosition, COLUMN_LINK_PHOTOS, value);
			if (linksString != null && !linksString.isEmpty()) {
				Collection<LinkData> links = toListLink(excelPosition, COLUMN_LINK_PHOTOS, linksString);
				retval.getPhotos().addAll(links);
			}
		}

		//
		// LINK_JOURNAL
		//
		index = header.get(COLUMN_LINK_JOURNAL);
		if (index != null) {
			excelPosition.setColumn(index);
			cell = row.getCell(index);
			String value = ExcelPoiUtils.readValueString(excelPosition, poiWorkbook, cell);
			Collection<String> linksString = toListString(excelPosition, COLUMN_LINK_JOURNAL, value);
			if (linksString != null && !linksString.isEmpty()) {
				Collection<LinkData> links = toListLink(excelPosition, COLUMN_LINK_JOURNAL, linksString);
				retval.getJournals().addAll(links);
			}
		}

		//
		// LINK_TRACES
		//
		index = header.get(COLUMN_LINK_TRACES);
		if (index != null) {
			excelPosition.setColumn(index);
			cell = row.getCell(index);
			String value = ExcelPoiUtils.readValueString(excelPosition, poiWorkbook, cell);
			Collection<String> linksString = toListString(excelPosition, COLUMN_LINK_TRACES, value);
			if (linksString != null && !linksString.isEmpty()) {
				Collection<LinkData> links = toListLink(excelPosition, COLUMN_LINK_TRACES, linksString);
				retval.getTraces().addAll(links);
			}
		}

		return retval;
	}

	private boolean isEmpty(Object value) {
		if (value == null) {
			return true;
		}

		if (value instanceof String) {
			return ((String) value).trim().isEmpty();
		}
		return false;
	}

	private Date toDate(Object value) {
		if (value instanceof Date) {
			return (Date) value;
		} else if (value instanceof String) {
			try {
				return DATE_FORMAT.parse((String) value);
			} catch (ParseException ex) {
				return null;
			}
		} else {
			return null;
		}
	}

	private Collection<String> toListString(ExcelPosition excelPosition, String column, String value) throws BaseException {
		Collection<String> retval = new ArrayList<String>();

		if (value == null) {
			return retval;
		}

		String[] parts = value.split(",");

		if (parts.length == 1 && isEmpty(parts[0])) {
			return retval;
		}

		for (int i = 0; i < parts.length; i++) {
			parts[i] = parts[i].trim();
			if (isEmpty(parts[i])) {
				throw new ExcelException(excelPosition, column + " value of list is empty");
			}
			retval.add(parts[i]);
		}

		return retval;
	}

	private Collection<LinkData> toListLink(ExcelPosition excelPosition, String column, Collection<String> linksString) throws BaseException {
		Collection<LinkData> retval = new ArrayList<LinkData>();
		if (linksString != null && !linksString.isEmpty()) {
			for (String value : linksString) {
				LinkData linkData = toLinkData(excelPosition, column, value);
				retval.add(linkData);
			}
		}
		return retval;
	}

	private LinkData toLinkData(ExcelPosition excelPosition, String column, String value) throws BaseException {
		String name = null;
		String urlString;
		URL url;
		if (value.startsWith("[")) {
			int pos = value.indexOf("]", 1);
			if (pos == -1) {
				throw new ExcelException(excelPosition, column + " Bad format of the [name] URL: " + value);
			}
			name = value.substring(1, pos).trim();
			urlString = value.substring(pos + 1).trim();
		} else {
			urlString = value.trim();
		}
		try {
			url = new URL(urlString);
		} catch (MalformedURLException ex) {
			throw new ExcelException(excelPosition, COLUMN_IMMAGE_MAP + " value is not a URL: " + value);
		}
		LinkData linkData = new LinkData(name, url);
		return linkData;
	}
}
