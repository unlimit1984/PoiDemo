import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by vladimir on 8/20/16.
 */
public class DocExample {

    public static void main(String[] args) throws IOException {
        XWPFDocument document= new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run=paragraph.createRun();
        run.setText("HelloWorld from Vladimir Vysokomornyi" +
                "during presentation \"Excel reports using Apache POI library\"." +
                "Wednesday-24-08-2016! ");

        FileOutputStream out = new FileOutputStream(new File("reports/file.docx"));
        document.write(out);
        out.close();
    }
}
