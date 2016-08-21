import org.apache.poi.xslf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Vladimir on 21.08.2016.
 */
public class PowerPointExample {
    public static void main(String[] args) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlideMaster slideMaster = ppt.getSlideMasters().get(0);
        XSLFSlideLayout titleLayout = slideMaster.getLayout(SlideLayout.TITLE);
        XSLFSlide slide1 = ppt.createSlide(titleLayout);
        XSLFTextShape title1 = slide1.getPlaceholder(0);
        title1.setText("My report");

        File file=new File("reports/power_point.pptx");
        FileOutputStream out = new FileOutputStream(file);
        ppt.write(out);
        out.close();    }
}