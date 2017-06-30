package kkr.travel.utils.excel.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import kkr.common.errors.BaseException;
import kkr.common.errors.TechnicalException;
import kkr.common.utils.UtilsFile;
import kkr.common.utils.UtilsResource;

public class PoiManipulator {
    private static final Logger LOG = Logger.getLogger(PoiManipulator.class);

    public static PoiWorkbook createWorkbook(File file)
            throws BaseException {
        LOG.trace("BEGIN: " + file.getAbsolutePath());
        try {
            PoiWorkbook poiWorkbook = new PoiWorkbook();
            poiWorkbook.setFile(file);

            Workbook workbook = new XSSFWorkbook();
            poiWorkbook.setWorkbook(workbook);

            FormulaEvaluator evaluator = new XSSFFormulaEvaluator(
                    (XSSFWorkbook) workbook);
            poiWorkbook.setEvaluator(evaluator);

            LOG.trace("OK");
            return poiWorkbook;
        } finally {
            LOG.trace("END: " + file.getAbsolutePath());
        }
    }

    public static PoiWorkbook openWorkbook(File file)
            throws BaseException {
        LOG.trace("BEGIN: " + file.getAbsolutePath());
        try {
            PoiWorkbook poiWorkbook = new PoiWorkbook();
            poiWorkbook.setFile(file);

            FileInputStream fileInputStream = null;
            try {
                fileInputStream = new FileInputStream(file);
                poiWorkbook.setInputStream(fileInputStream);

                Workbook workbook = WorkbookFactory.create(fileInputStream);
                poiWorkbook.setWorkbook(workbook);

                FormulaEvaluator evaluator = new XSSFFormulaEvaluator(
                        (XSSFWorkbook) workbook);
                poiWorkbook.setEvaluator(evaluator);

            } catch (InvalidFormatException ex) {
                throw new TechnicalException("Bad format of the excel file: "
                        + file.getAbsolutePath(), ex);
            } catch (FileNotFoundException ex) {
                throw new TechnicalException("The file does not exist: "
                        + file.getAbsolutePath(), ex);
            } catch (IOException ex) {
                throw new TechnicalException("Cannot read the excel file: "
                        + file.getAbsolutePath(), ex);
            }

            LOG.trace("OK");
            return poiWorkbook;
        } finally {
            LOG.trace("END: " + file.getAbsolutePath());
        }
    }

    public static void closeWorkbook(PoiWorkbook poiWorkbook)
            throws BaseException {
        LOG.trace("BEGIN");
        try {
            if (poiWorkbook != null) {
                try {
                    if (poiWorkbook.getInputStream() != null) {
                        poiWorkbook.getInputStream().close();
                        poiWorkbook.setInputStream(null);
                    }
                    LOG.trace("OK");
                } catch (IOException ex) {
                    throw new TechnicalException("Cannot close the workbook: "
                            + poiWorkbook.getFile().getAbsolutePath(), ex);
                } finally {
                    UtilsResource.closeResource(poiWorkbook.getInputStream());
                }
            }
            LOG.trace("OK");
        } finally {
            LOG.trace("END");
        }
    }

    public static void saveWorkbook(PoiWorkbook poiWorkbook)
            throws BaseException {
        LOG.trace("BEGIN: " + poiWorkbook.getFile().getAbsolutePath());
        try {
            saveWorkbook(poiWorkbook, poiWorkbook.getFile());
            LOG.trace("OK");
        } finally {
            LOG.trace("END: " + poiWorkbook.getFile().getAbsolutePath());
        }
    }

    public static void saveWorkbook(PoiWorkbook poiWorkbook, File file)
            throws BaseException {
        LOG.trace("BEGIN: " + file.getAbsolutePath());
        try {
            Workbook workbook = poiWorkbook.getWorkbook();
            UtilsFile.createFileDirectory(file);

            FileOutputStream outputStream = null;
            try {
                outputStream = new FileOutputStream(file);
                workbook.write(outputStream);
                outputStream.close();
                outputStream = null;
            } catch (IOException ex) {
                throw new TechnicalException("Cannot write the file: "
                        + file.getAbsolutePath(), ex);
            } finally {
                UtilsResource.closeResource(outputStream);
            }

            LOG.trace("OK");
        } finally {
            LOG.trace("END: " + file.getAbsolutePath());
        }
    }
}
