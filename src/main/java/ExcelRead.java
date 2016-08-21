import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

/**
 * Created by Vladimir on 21.08.2016.
 */
public class ExcelRead {

    public static SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd");

    public static void main(String[] args) throws IOException {

        FileInputStream fileIn = new FileInputStream("C:/ALL/temp/read.xls");
        Workbook wb = new HSSFWorkbook(fileIn);

        for(int i=0;i<wb.getNumberOfSheets();i++){
            Sheet sheet = wb.getSheetAt(i);
            for(Row row : sheet){
                for(Cell cell : row){
                    CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                    System.out.print(cellRef.formatAsString());
                    System.out.print(" - ");
                    System.out.print("row=" + cellRef.getRow()+" - ");
                    System.out.print("column=" + cellRef.getCol()+" - ");

                    System.out.println(getCellText(cell));
                }
            }
        }
        fileIn.close();


    }
    public static String getCellText(Cell cell){

        String result="";

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = sdf.format(cell.getDateCellValue());
                } else {
                    result = Double.toString(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                result = Boolean.toString(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                result = cell.getCellFormula().toString();
                break;
            default:
                break;
        }
        return result;
    }

}
