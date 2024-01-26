package com.sample.excelOperation;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter implements Runnable {
    private String filePath;

    public ExcelWriter(String filePath) {
        this.filePath = filePath;
    }

    @Override
    public void run() {
        try (Workbook workbook = new XSSFWorkbook()) { // XSSFWorkbook for .xlsx format

            Sheet sheet = workbook.createSheet("Sheet1");

            // Example data for writing
            String[][] data = {
                    {"Name", "Age", "City"},
                    {"John", "25", "New York"},
                    {"Alice", "30", "San Francisco"}
            };

            int rowNum = 0;
            for (String[] rowData : data) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (String value : rowData) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(value);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(new File(filePath))) {
                workbook.write(fos);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
