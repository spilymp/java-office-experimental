package poi;

import org.apache.poi.xslf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * Example to add notes to a slide using Apache POI (XSLF).
 * Used apache poi version: 3.12
 *
 * @author spilymp
 */
public class SlideNotes {

    private static final String path = System.getProperty("user.home") + "/documents/";

    public static void main(String[] args) {

        // create a new empty slide show
        XMLSlideShow pptx = new XMLSlideShow();

        // add first slide
        XSLFSlide slide = pptx.createSlide();

        // get or create notes
        XSLFNotes note = pptx.getNotesSlide(slide);

        // insert text
        for (XSLFTextShape shape : note.getPlaceholders()) {
            if (shape.getTextType() == Placeholder.BODY) {
                shape.setText("Lorem ipsum dolor sit ...");
                break;
            }
        }

        // save file
        try {
            OutputStream os = new FileOutputStream(path + "SlideNotes.pptx");
            pptx.write(os);
            os.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
