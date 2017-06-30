package kkr.travel.utils.excel.poi;

import java.io.File;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

public class PoiWorkbook {
    private Workbook workbook;
    private InputStream inputStream;
    private File file;
    private FormulaEvaluator evaluator;

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }

    public File getFile() {
        return file;
    }

    public void setFile(File file) {
        this.file = file;
    }

    public FormulaEvaluator getEvaluator() {
        return evaluator;
    }

    public void setEvaluator(FormulaEvaluator evaluator) {
        this.evaluator = evaluator;
    }
}
