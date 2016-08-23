import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Vladimir on 24.08.2016.
 */
public class ExcelShape {

    public static void main(String[] args) throws IOException {

        Workbook wb = new HSSFWorkbook();
        Sheet sheet=wb.createSheet("shapes");

        HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();

        HSSFClientAnchor anchor = new HSSFClientAnchor();
        anchor.setCol1(2);
        anchor.setRow1(2);
        anchor.setCol2(10);
        anchor.setRow2(10);

        HSSFSimpleShape shape = patriarch.createSimpleShape(anchor);
        shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
        shape.setLineStyleColor(255, 0, 0);
        shape.setLineWidth(HSSFSimpleShape.LINEWIDTH_ONE_PT*3);
        shape.setLineStyle(HSSFSimpleShape.LINESTYLE_DASHDOTGEL);
        shape.setFillColor(0, 0, 255);;


        FileOutputStream fos = new FileOutputStream("reports/shape.xls");
        wb.write(fos);
        fos.close();
    }
}
