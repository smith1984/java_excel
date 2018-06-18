package ru.smith.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelStyles {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet =workbook.createSheet("Формулы");

        Row row = sheet.createRow(0);
        CellStyle style = workbook.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        //style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        Font font = workbook.createFont();
        font.setFontName("Courier New");
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        font.setStrikeout(true);
        font.setUnderline(Font.U_SINGLE);
        font.setColor(IndexedColors.RED.getIndex());
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.DOUBLE);
        style.setBottomBorderColor(IndexedColors.YELLOW.getIndex());
        Cell cell = row.createCell(0);
        cell.setCellStyle(style);
        cell.setCellValue("Привет");

        FileOutputStream fos = new FileOutputStream("styles.xls");
        workbook.write(fos);
        fos.close();
        workbook.close();
    }
}
