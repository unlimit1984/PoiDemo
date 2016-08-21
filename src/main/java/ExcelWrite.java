import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Created by Vladimir on 21.08.2016.
 */
public class ExcelWrite {
    public static void main(String[] args) throws IOException {
        Workbook wb = new HSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();

        Sheet sheet0=wb.createSheet("Publishers");
        Row row = sheet0.createRow(3);
        Cell cell = row.createCell(4);
        cell.setCellValue("O'Reilly");

        cell = row.createCell(5);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-mm-yyyy"));
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);
        sheet0.autoSizeColumn(5);

        Sheet sheet1=wb.createSheet("Novels");
        Row row1 = sheet1.createRow(0);
        Cell cell1 = row1.createCell(0);
        cell1.setCellValue("War and Peace");

        Row row2 = sheet1.createRow(1);
        Cell cell2 = row2.createCell(3);
        cell2.setCellValue("Flowers for Algernon");

        Sheet sheet2=wb.createSheet("Authors");
        Sheet sheet3=wb.createSheet(WorkbookUtil.createSafeSheetName("a[b]c*d/e\\f"));

        FileOutputStream fos = new FileOutputStream("reports/publishers.xls");
        wb.write(fos);
        fos.close();

    }
}
