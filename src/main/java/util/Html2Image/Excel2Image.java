import examples.toHTML.ToHtml;
import gui.ava.html.image.generator.HtmlImageGenerator;
import org.apache.log4j.Logger;

import java.awt.image.BufferedImage;
import java.io.*;

/**
 * Convert an excel file to an image. First step is to transform to HTML. Thereafter the HTML is transformed to an
 * image with <a href="https://code.google.com/p/java-html2image/">Html2Image</a>.
 *
 * @author spilymp
 */
public class Excel2Image {

    private static final String s = File.separator;
    private static final String path = System.getProperty("user.home") + s + "Documents" + s + "Tests" + s + "ApachePOI" + s;

    // html string
    private static String htmlString = "";

    // logger
    private static final Logger log = Logger.getLogger(Excel2Image.class);

    public static void main(String[] args) {
        Excel2Image converter = new Excel2Image();
        htmlString = converter.transformExcel2HTML();
        converter.transformHTML2Image();
    }

    /**
     * Transform an excel file to HTML.
     *
     * @return the HTML file as string.
     */
    public String transformExcel2HTML() {
        // load excel file
        FileInputStream xlsx = null;
        try {
            xlsx = new FileInputStream(path + "resources" + s + "xlsx" + s + "exampleTable.xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }

        // if excel file could no be loaded, return
        if (xlsx == null) {
            log.error("Excel file is null, function returned.");
            return "";
        }

        StringWriter writer = new StringWriter();
        ToHtml html = null;

        try {
            // TODO: Change path to standard css file!
            html = ToHtml.create(xlsx, writer);
        } catch (IOException e) {
            e.printStackTrace();
        }

        if (html != null) {
            html.setCompleteHTML(true);
            try {
                html.printPage();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            log.error("HTML is null.");
        }

        // write html to file, only for debugging
        FileWriter fw = null;
        try {
            fw = new FileWriter(path + "html.xhtml");
            fw.write(writer.toString());
            writer.close();
            fw.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // return html
        return writer.toString();
    }

    /**
     * Transform HTML to image
     */
    public BufferedImage transformHTML2Image() {
        HtmlImageGenerator imageGenerator = new HtmlImageGenerator();
        imageGenerator.loadHtml(htmlString);
        // save image for debugging
        imageGenerator.saveAsImage(path + "excel2image.png");
        // return image
        return imageGenerator.getBufferedImage();
    }
}
