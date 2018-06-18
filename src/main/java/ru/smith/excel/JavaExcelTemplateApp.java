package ru.smith.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelTemplateApp {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet =workbook.createSheet("Лист");

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);


        FileOutputStream fos = new FileOutputStream("template.xls");
        workbook.write(fos);
        fos.close();
        workbook.close();
    }
}