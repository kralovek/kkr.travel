package kkr.travel.utils.excel;

import kkr.common.errors.ConfigurationException;

public class ExcelUtils implements ExcelConstants {

    private static ExcelUtils excelUtils = new ExcelUtils();

    public static ExcelUtils getInstance() {
        return excelUtils;
    }

    public ExcelIdCell adaptAndCheckCellId(final String pCellId,
            final String pName) throws ConfigurationException {
        if (pCellId == null) {
            return null;
        }
        int i;
        for (i = 0; i < pCellId.length()
                && Character.isLetter(pCellId.charAt(i)); i++) {
        }
        final String idColumnCell = pCellId.substring(0, i);
        final String idRowCell = pCellId.substring(i);

        final Integer idColumnExcelMin = columnToIndex(EXCEL_MIN_COLUMN);
        final Integer idColumnExcelMax = columnToIndex(EXCEL_MAX_COLUMN);
        final Integer columnIntId = columnToIndex(idColumnCell);
        Integer rowIntId;
        try {
            rowIntId = Integer.parseInt(idRowCell);
        } catch (final NumberFormatException ex) {
            rowIntId = null;
        }

        if (columnIntId == null || rowIntId == null
                || //
                columnIntId < idColumnExcelMin
                || columnIntId > idColumnExcelMax || //
                rowIntId < EXCEL_MIN_ROW || rowIntId > EXCEL_MAX_ROW) {
            throw new ConfigurationException(getClass().getSimpleName() + ": "
                    + pName + " must be form the range: <" + EXCEL_MIN_COLUMN
                    + EXCEL_MIN_ROW + " - " + EXCEL_MAX_COLUMN + EXCEL_MAX_ROW
                    + ">");
        }
        final ExcelIdCell idCell = new ExcelIdCell();
        idCell.setColumn(columnIntId);
        idCell.setRow(rowIntId - 1);

        return idCell;
    }

    public Integer adaptAndCheckRowId(final Integer pRowId, final String pName)
            throws ConfigurationException {
        if (pRowId == null) {
            return null;
        }
        if (pRowId < EXCEL_MIN_ROW || pRowId > EXCEL_MAX_ROW) {
            throw new ConfigurationException(getClass().getSimpleName() + ": "
                    + pName + " must be form the range: <" + EXCEL_MIN_ROW
                    + " - " + EXCEL_MAX_ROW + ">");
        }
        return pRowId - 1;
    }

    public Integer adaptAndCheckColumnId(final String pColumnId,
            final String pName) throws ConfigurationException {
        if (isEmpty(pColumnId)) {
            return null;
        }
        final Integer idColumnExcelMin = columnToIndex(EXCEL_MIN_COLUMN);
        final Integer idColumnExcelMax = columnToIndex(EXCEL_MAX_COLUMN);
        final Integer columnIntId = columnToIndex(pColumnId);

        if (columnIntId == null || columnIntId < idColumnExcelMin
                || columnIntId > idColumnExcelMax) {
            throw new ConfigurationException(getClass().getSimpleName() + ": "
                    + pName + " must be from the range: <" + EXCEL_MIN_COLUMN
                    + " - " + EXCEL_MAX_COLUMN + ">");
        }
        return columnIntId;
    }

    public Integer columnToIndex(final String pColumn) {
        if (pColumn.length() == 0 || pColumn.length() > 2) {
            return null;
        }
        final String column = pColumn.toUpperCase();

        if (!Character.isLetter(column.charAt(0)) || column.length() == 2
                && !Character.isLetter(pColumn.charAt(1))) {
            return null;
        }

        int a = ((int) column.charAt(0)) - ((int) 'A');
        if (column.length() == 2) {
            int b = ((int) column.charAt(1)) - ((int) 'A');
            return (a + 1) * ('Z' - 'A') + b;
        } else {
            return a;
        }
    }

    protected boolean isEmpty(final String pValue) {
        return pValue == null || pValue.isEmpty();
    }

    public Integer checkColumn(String parameter, String column, boolean mandatory) throws ConfigurationException {
        if (column == null) {
            if (mandatory) {
                throw new ConfigurationException("The parameter '" + parameter + "' is not configured");
            } else {
                return null;
            }
        } else {
            return ExcelUtils.getInstance().adaptAndCheckColumnId(column, "Parameter '" + parameter + "'");
        }
    }

    public Integer checkRow(String parameter, Integer row, boolean mandatory) throws ConfigurationException {
        if (row == null) {
            if (mandatory) {
                throw new ConfigurationException("The parameter '" + parameter + "' is not configured");
            } else {
                return null;
            }
        } else {
            return ExcelUtils.getInstance().adaptAndCheckRowId(row, "Parameter '" + parameter + "'");
        }
    }
}
