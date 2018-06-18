package ru.smith.excel;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class JavaExcelPictureApp {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet =workbook.createSheet("Лист");

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);

        HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
        HSSFClientAnchor anchor = new HSSFClientAnchor();
        anchor.setCol1(2);
        anchor.setCol2(10);
        anchor.setRow1(2);
        anchor.setRow2(10);

        HSSFSimpleShape simpleShape = patriarch.createSimpleShape(anchor);
        simpleShape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);
        simpleShape.setLineStyleColor(255, 0, 0);
        simpleShape.setLineWidth(HSSFSimpleShape.LINEWIDTH_ONE_PT*3);
        simpleShape.setLineStyle(HSSFSimpleShape.LINESTYLE_DASHDOTGEL);



        FileOutputStream fos = new FileOutputStream("picture.xls");
        workbook.write(fos);
        fos.close();
        workbook.close();
    }
}
