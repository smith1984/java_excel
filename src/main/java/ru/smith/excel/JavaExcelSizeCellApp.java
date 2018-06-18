package ru.smith.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelSizeCellApp {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet =workbook.createSheet("Лист");
        sheet.setColumnWidth(3, 10000);

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Новая ячейка");
        sheet.autoSizeColumn(0);
        row.setHeightInPoints(20);
        sheet.addMergedRegion(new CellRangeAddress(0,3,0,4));

        FileOutputStream fos = new FileOutputStream("SizeCell.xls");
        workbook.write(fos);
        fos.close();
        workbook.close();
    }
}
