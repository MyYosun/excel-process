package com.zyf.excelprocess;

import org.apache.poi.hssf.usermodel.*;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

@Component
public class OutExcelUtil {

    @Value("${excel.output.path}")
    private static String outPath;

    @Value("${excel.output.name}")
    private static String outName;


    /**
     * 导出Excel
     *
     * @param sheetName sheet名称
     * @param title     标题
     * @param values    内容
     * @param wb        HSSFWorkbook对象
     * @return
     */
    public HSSFWorkbook getHSSFWorkbook(String sheetName, String[] title, List<Map> values, HSSFWorkbook wb) {

        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        if (wb == null) {
            wb = new HSSFWorkbook();
        }

        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(sheetName);

        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        HSSFRow row = sheet.createRow(0);

        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式

        //声明列对象
        HSSFCell cell = null;

        //创建标题
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }

        //创建内容
        for (int i = 0; i < values.size(); i++) {
            row = sheet.createRow(i + 1);
            Map singleItem = values.get(i);
            row.createCell(0).setCellValue(singleItem.get("name").toString());
            row.createCell(1).setCellValue(singleItem.get("in_time").toString());
            row.createCell(2).setCellValue(singleItem.get("out_time").toString());
        }

        return wb;
    }

    public void processOutputData(HSSFWorkbook hssfWorkbook, Date date, List<Map> data) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-SS");
        String name = simpleDateFormat.format(date);
        String[] title = new String[]{"人员", "进场", "出场"};
        getHSSFWorkbook(name, title, data, hssfWorkbook);
    }

    public void outputExcel(HSSFWorkbook hssfWorkbook) {
        String path = outPath;
        File outDir = new File(path);
        if (!outDir.exists()) {
            outDir.mkdir();
        }

        try {
            hssfWorkbook.write(new File(path + File.separator + "outName"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
