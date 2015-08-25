package util.Html2Image;import gui.ava.html.image.generator.HtmlImageGenerator;

import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @author spilymp
 */
public class Html2Image {

    private static final String path = "C:/Users/local_user/IdeaProjects/experimental/src/main/resources/";

    public static void main(String[] args) {
        new Html2Image().html2Image();
    }

    public void html2Image() {
        HtmlImageGenerator imageGenerator = new HtmlImageGenerator();
        String htmlString = readFile(path + "test.xhtml", StandardCharsets.UTF_8);
        imageGenerator.loadHtml(htmlString);
        imageGenerator.saveAsImage(System.getProperty("user.home") + File.separator + "hello-world.png");
    }

    static String readFile(String path, Charset encoding) {
        byte[] encoded = new byte[0];
        try {
            encoded = Files.readAllBytes(Paths.get(path));
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new String(encoded, encoding);
    }
}
