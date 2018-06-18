package ru.smith.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelFormulaApp {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet =workbook.createSheet("Формулы");

        Row row = sheet.createRow(0);

        Cell cell = row.createCell(0);
        cell.setCellValue(2);

        cell = row.createCell(1);
        cell.setCellValue(7);

        cell = row.createCell(2);
        cell.setCellFormula("A1+B1");

        for (int i =1; i <10; i++){
        row =sheet.createRow(i);
        cell = row.createCell(0);
        cell.setCellValue(i*i);
        }
        row = sheet.createRow(10);
        cell = row.createCell(0);
        cell.setCellFormula("SUM(A1:A10)");


        FileOutputStream fos = new FileOutputStream("formula.xls");
        workbook.write(fos);
        fos.close();
        workbook.close();
    }
}
