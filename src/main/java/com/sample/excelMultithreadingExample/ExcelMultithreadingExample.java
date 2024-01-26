package com.sample.excelMultithreadingExample;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class ExcelMultithreadingExample {

    public static void main(String[] args) {
        // Specify the Excel file path
        String filePath = "D:\\excel-practice\\November-clearance.xlsx";
        //"path/to/your/excel/file.xlsx";

        // Specify the number of threads
        int numberOfThreads = 4;

        // Create a thread pool with the desired number of threads
        ExecutorService executorService = Executors.newFixedThreadPool(numberOfThreads);

        try {
            // Open the Excel file
            FileInputStream fileInputStream = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(fileInputStream);

            // Get the number of sheets in the Excel workbook
            int numberOfSheets = workbook.getNumberOfSheets();

            // Iterate through each sheet and submit it to the thread pool for processing
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                executorService.submit(new ExcelSheetProcessor(sheet));
            }

            // Shutdown the thread pool
            executorService.shutdown();

            // Wait for all threads to finish
            while (!executorService.isTerminated()) {
                // Wait
            }

            // Close the workbook and file input stream
            workbook.close();
            fileInputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

class ExcelSheetProcessor implements Runnable {
    private Sheet sheet;

    public ExcelSheetProcessor(Sheet sheet) {
        this.sheet = sheet;
    }

    @Override
    public void run() {
        // Process the sheet (e.g., read or write data)
        // Add your sheet processing logic here
        System.out.println("Processing sheet: " + sheet.getSheetName());
    }
}
