package ru.smith.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelWriteApp {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet_0 = workbook.createSheet("Издатели");

        Row row = sheet_0.createRow(3);
        Cell cell = row.createCell(4);
        cell.setCellValue("O'Railly");

        Sheet sheet_1 = workbook.createSheet("Книги");
        Row row_1 = sheet_1.createRow(0);
        Cell cell_1 = row_1.createCell(0);
        cell_1.setCellValue("Война и мир");
        Row row_2 = sheet_1.createRow(1);
        Cell cell_2 = row_2.createCell(0);
        cell_2.setCellValue("ПетрI");

        Sheet sheet_2 = workbook.createSheet("Авторы");
        Sheet sheet_3 = workbook.createSheet(WorkbookUtil.createSafeSheetName("adsfgsd@#%$&^(*&^"));

        FileOutputStream fos = new FileOutputStream("myExcelBook.xls");

        workbook.write(fos);

        fos.close();
    }

}
