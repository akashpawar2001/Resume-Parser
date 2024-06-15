package com.resumeextractor.extractor.util;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelWriter {
    private Workbook workbook;
    private Sheet sheet;
    private int rowNum;

    public ExcelWriter() {
        this.workbook = new XSSFWorkbook();
        this.sheet = workbook.createSheet("Resumes");
        this.rowNum = 0;
        createHeaderRow();
    }

    private void createHeaderRow() {
        Row headerRow = sheet.createRow(rowNum++);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Email");
        headerRow.createCell(2).setCellValue("Phone");
        headerRow.createCell(3).setCellValue("Education");
        headerRow.createCell(4).setCellValue("Current Company");
    }

    public void addRow(String name,String emails, String phone,String education,String currentCompany) {
        Row row = sheet.createRow(rowNum++);
        row.createCell(0).setCellValue(name);
        row.createCell(1).setCellValue(emails);
        row.createCell(2).setCellValue(phone);
        row.createCell(3).setCellValue(education);
        row.createCell(4).setCellValue(currentCompany);
    }

    public void saveToFile(String filePath) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
        } finally {
            workbook.close();
        }
    }
}
