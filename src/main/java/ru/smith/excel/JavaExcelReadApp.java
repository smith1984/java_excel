package ru.smith.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

public class JavaExcelReadApp {

    public static void main(String[] args) throws IOException {

        FileInputStream fis = new FileInputStream("FileExcelToRead.xls");
        Workbook workbook = new HSSFWorkbook(fis);

        StringBuilder text;

        for (Row row : workbook.getSheetAt(0)) {
            for (Cell cell : row) {
                text = new StringBuilder();
                text.append(getNameCell(cell));
                text.append(" - ");
                text.append(getCellText(cell));
                System.out.println(cell.getCellStyle().getIndex());
                System.out.println(text);
            }
        }

        fis.close();

    }

    public static String getNameCell(Cell cell) {
        CellReference cellReference = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
        return cellReference.formatAsString();

    }

    public static String getCellText(Cell cell) {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
        DataFormatter formatter = new DataFormatter();
        if (cell.getCellStyle().getIndex() == 22)
            formatter.addFormat(cell.getCellStyle().getDataFormatString(), sdf);
        return formatter.formatCellValue(cell);

    }
}
