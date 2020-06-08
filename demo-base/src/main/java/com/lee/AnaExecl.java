package com.lee;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Hello world!
 *
 */
public class AnaExecl
{
    public static void main( String[] args ) throws IOException {
        File file = new File("demo-base/Test.xls");
        System.out.println(file.getAbsolutePath());
        if (!file.exists()){
            file.createNewFile();
        }
        FileOutputStream stream = new FileOutputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("测试");
        HSSFRow row = sheet.createRow(1);
        row.createCell(0).setCellValue("Hello0000000000000000");
        row.createCell(1).setCellValue("World000000000000000");
        CellRangeAddress addresses =new CellRangeAddress(3,3,4,7);
        sheet.addMergedRegion(addresses);
        sheet.autoSizeColumn(0);
        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFont(font);
        row.getCell(0).setCellStyle(style);

        workbook.setActiveSheet(0);
        workbook.write(stream);
        stream.close();
    }
}
