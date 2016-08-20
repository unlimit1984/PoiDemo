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
        //Blank Document
        XWPFDocument document= new XWPFDocument();
        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(new File("file.docx"));

        //create Paragraph
        XWPFParagraph paragraph = document.createParagraph();

        XWPFRun run=paragraph.createRun();

        run.setText("HelloWorld from Vladimir Vysokomornyi during presentation \"Excel reports using Apache POI library\". Wednesday-24-08-2016! ");

        document.write(out);
        out.close();
        System.out.println("createparagraph.docx written successfully");

    }



}
