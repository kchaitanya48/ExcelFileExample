package com.sample.excelOperation;

import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class ExcelMultithreadingExample {
    public static void main(String[] args) {
        String inputFile = "D:\\\\excel-practice\\input.xlsx";
        String outputFile = "D:\\\\excel-practice\\output.xlsx";

        ExecutorService executorService = Executors.newFixedThreadPool(2);

        // Create instances of your reader and writer
        ExcelReader excelReader = new ExcelReader(inputFile);
        ExcelWriter excelWriter = new ExcelWriter(outputFile);

        // Submit threads to the executor
        executorService.submit(excelReader);
        executorService.submit(excelWriter);

        // Shutdown the executor when tasks are complete
        executorService.shutdown();
    }
}
