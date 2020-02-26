package ru.ak.model;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.Data;

/**
 * Описание Excel-файла
 */
@Data
public class ExcelFile {

    /** Полное имя файла */
    private String fileName;

    /** Книга */
    private Workbook workbook;

    /** Это XLSX-файл */
    private boolean is2007;

    public ExcelFile(String fileName) throws IOException {
        setFileName(fileName);
    }

    public void setFileName(String fileName) throws IOException {
        this.fileName = fileName;

        String extension = fileName.substring(fileName.lastIndexOf("."));
        is2007 = extension.toLowerCase().equalsIgnoreCase(".xlsx");

        FileInputStream fis = new FileInputStream(fileName);

        this.workbook = is2007 ? new XSSFWorkbook(fis) : new HSSFWorkbook(fis);
    }

    public Workbook getWorkbook() {
        return this.workbook;
    }

    public String getFileName() {
        return this.fileName;
    }
}