package com.sample.excelOperation;

import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class ExcelMultithreadingExampleTest {

    public static void main(String[] args) {
        String inputFile = "D:\\\\excel-practice\\input.xlsx";
        String outputFile = "D:\\\\excel-practice\\output.xlsx";

   

        // Create instances of your reader and writer
        ExcelReader excelReader = new ExcelReader(inputFile);
        ExcelWriter excelWriter = new ExcelWriter(outputFile);

        // Submit threads to the executor
        
        
        Thread read =new Thread(excelReader);
        Thread write =new Thread(excelWriter);
        
    read.start();
    write.start();

    }


}
