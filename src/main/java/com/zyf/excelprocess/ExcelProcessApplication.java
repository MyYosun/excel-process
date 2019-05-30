package com.zyf.excelprocess;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.InputStream;

import static com.sun.org.apache.xerces.internal.utils.SecuritySupport.getResourceAsStream;

@SpringBootApplication
public class ExcelProcessApplication {


    public static void main(String[] args) {
        InputStream in = getResourceAsStream("person.xlsx");

        try {
            Workbook workbook = WorkbookFactory.create(in);
            Sheet sheet = workbook.getSheetAt(0);
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                System.out.println(row.getCell(0).toString());
            }

        } catch (Exception e) {

        }


        System.out.println("running...");
        SpringApplication.run(ExcelProcessApplication.class, args);
    }

}
