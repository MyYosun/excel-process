package com.zyf.excelprocess;

import com.ling.Person;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static com.sun.org.apache.xerces.internal.utils.SecuritySupport.getResourceAsStream;

@SpringBootApplication
public class ExcelProcessApplication {


    public static void main(String[] args) {

//        List<Person> personList = getPersonList();
//        System.out.println(personList.size());

        getInAndOut();

        // test
        // test


        SpringApplication.run(ExcelProcessApplication.class, args);
    }


    // 从花名册Excel中读取人员信息
    public static List<Person> getPersonList() {
        InputStream in = getResourceAsStream("person.xlsx");
        List<Person> personList = new ArrayList<>();
        Person person = new Person();
        try {
            Workbook workbook = WorkbookFactory.create(in);
            Sheet sheet = workbook.getSheetAt(0);
            for (int rowNum = 5; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                String id = row.getCell(0).toString();
                if (id != null && !"".equals(id)) {
                    String name = row.getCell(1).toString();
                    person.setId(id);
                    person.setName(name);
                    person.setStatus(0);
                    personList.add(person);
                }
            }
        } catch (Exception e) {
            System.out.println("---------程序异常----------");
            e.printStackTrace();
            System.out.println("---------程序异常----------");
        }
        return personList;
    }



    // 获取每日各时段Excel
    public static void getInAndOut() {
        InputStream in = getResourceAsStream("inandout.xlsx");
        try {
            Workbook workbook = WorkbookFactory.create(in);
            Sheet sheet = workbook.getSheetAt(0);
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                Date date = row.getCell(0).getDateCellValue();
                System.out.println("日期: " + date);
            }
        } catch (Exception e) {
            System.out.println("---------程序异常----------");
            e.printStackTrace();
            System.out.println("---------程序异常----------");
        }
    }

}