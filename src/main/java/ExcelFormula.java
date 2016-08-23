import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Vladimir on 23.08.2016.
 */
public class ExcelFormula {
    public static void main(String[] args) throws IOException {

        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Formulas");

        Row row = sheet.createRow(0);

        Cell cell0 = row.createCell(0);
        cell0.setCellValue(2);

        Cell cell1 = row.createCell(1);
        cell1.setCellValue(7);

        Cell cell2 = row.createCell(2);
        cell2.setCellFormula("A1*B1");


        Row row3 = sheet.createRow(3);
        Cell cell3 = row3.createCell(0);
        cell3.setCellValue(1);

        Row row4 = sheet.createRow(4);
        Cell cell4 = row4.createCell(0);
        cell4.setCellValue(2);

        Row row5 = sheet.createRow(5);
        Cell cell5 = row5.createCell(0);
        cell5.setCellValue(3);

        Row row6 = sheet.createRow(6);
        Cell cell6 = row6.createCell(0);
        cell6.setCellValue(4);

        Row row7 = sheet.createRow(7);
        Cell cell7 = row7.createCell(0);
        cell7.setCellFormula("SUM(A4:A7)");


        FileOutputStream fos = new FileOutputStream("reports/Formulas.xls");
        wb.write(fos);
        fos.close();
        wb.close();
    }
}
