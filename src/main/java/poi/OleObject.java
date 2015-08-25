package poi;

import org.apache.log4j.Logger;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFRelation;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectData;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPoint2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTransform2D;
import org.openxmlformats.schemas.presentationml.x2006.main.CTGraphicalObjectFrame;

import java.io.*;

/**
 * Example of how to add an excel-file (ole-object) to a slide using Apache POI (XSLF).
 * Used apache poi version: 3.12
 * <p/>
 * Many thanks to the author of the following post: <a href="http://pastebin.com/pu48nbW2">pastebin.com/pu48nbW2</a>
 *
 * @author spilymp
 */
public class OleObject {

    // working path
    private static final String path = System.getProperty("user.home") + "/documents/";

    // logger
    private static final Logger log = Logger.getLogger(OleObject.class);

    public static void main(String[] args) {
        //create a new empty slide show
        XMLSlideShow pptx = new XMLSlideShow();

        //add a slide
        XSLFSlide slide = pptx.createSlide();

        embed(slide);

        // save
        try {
            OutputStream os = new FileOutputStream(path + "ole.pptx");
            pptx.write(os);
            os.close();
        } catch (FileNotFoundException e) {
            log.error("Could not save pptx. ", e);
        } catch (IOException e) {
            log.error("Could not close outputstream. ", e);
        }
    }

    public static void embed(XSLFSlide slide) {

        // path of excel file
        String pathToXlsx = path + "resources" + File.separator + "xlsx" + File.separator + "exampleTable.xlsx";
        // path of preview image
        String pathToEmf = path + "resources" + File.separator + "img" + File.separator + "example.emf";

        // add xslt file-part to presentation
        String xlsId = null;
        try {
            xlsId = addPart(
                    slide,
                    pathToXlsx,
                    "/ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");
        } catch (Exception e) {
            log.error("Could not add excel-part to pptx. ", e);
        }

        // add preview image-part to presentation
        String imgId = null;
        try {
            imgId = addPart(
                    slide,
                    pathToEmf,
                    "/ppt/media/image6.emf",
                    "imagex/x-emf",
                    XSLFRelation.IMAGES.getRelation());
        } catch (Exception e) {
            log.error("Could not add image-part to pptx. ", e);
        }

        if (xlsId == null || imgId == null) {
            log.error("Error, imgId or xlsId were not added correctly.");
            return;
        }

        // add shape to slide
        CTGraphicalObjectFrame ch = slide.getXmlObject().getCSld().getSpTree().addNewGraphicFrame();
        try {
            insertXObject(ch, XmlObject.Factory.parse(shapeTmpl));
        } catch (XmlException e) {
            log.error("Could not add shape to pptx. ", e);
        }

        CTGraphicalObjectData gr = ch.addNewGraphic().addNewGraphicData();
        gr.setUri("http://schemas.openxmlformats.org/presentationml/2006/ole");

        CTTransform2D ctTransform2D = ch.addNewXfrm();
        CTPoint2D ctPoint2D = ctTransform2D.addNewOff();

        /**
         * define anchor
         * to use pixel/points use {@link org.apache.poi.util.Units}
         */
        ctPoint2D.setX(3352800);
        ctPoint2D.setY(3886200);
        CTPositiveSize2D ctPositiveSize2D = ctTransform2D.addNewExt();
        ctPositiveSize2D.setCx(2611438);
        ctPositiveSize2D.setCy(1506537);

        /**
         * set image-id and excel-id
         * if you changed the anchor above you have to change the template too {@link xlsEmbedTmpl}
         */
        String altCtx = xlsEmbedTmpl
                .replace("@xlsId", xlsId)
                .replace("@imgId", imgId);

        try {
            insertXObject(gr, XmlObject.Factory.parse(altCtx));
        } catch (XmlException e) {
            log.error("Could not add ole-object to pptx. ", e);
        }
    }

    private static void insertXObject(XmlObject root, XmlObject child) {
        XmlCursor rootCursor = root.newCursor();
        rootCursor.toEndToken();

        XmlCursor anyFooCursor = child.newCursor();
        anyFooCursor.toNextToken();
        anyFooCursor.moveXml(rootCursor);
    }

    private static String addPart(XSLFSlide slide, String inputFile, String zipPath, String contentType, String relationType) throws Exception {
        PackagePartName ppn = PackagingURIHelper.createPartName(zipPath);

        PackagePart xlsPart = slide.getPackagePart().getPackage().createPart(ppn, contentType);
        InputStream is = new FileInputStream(inputFile);
        OutputStream os = xlsPart.getOutputStream();
        IOUtils.copy(is, os);
        os.close();
        is.close();

        PackageRelationship rel = slide.getPackagePart().addRelationship(
                ppn, TargetMode.INTERNAL, relationType);

        return rel.getId();
    }

    private final static String shapeTmpl =
            "<p:nvGraphicFramePr " +
                    "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                    "<p:cNvPr id=\"5\" name=\"Objekt 4\"/>" +
                    "<p:cNvGraphicFramePr>" +
                    "<a:graphicFrameLocks noChangeAspect=\"1\"/>" +
                    "</p:cNvGraphicFramePr>" +
                    "<p:nvPr>" +
                    "<p:extLst>" +
                    "<p:ext uri=\"{D42A27DB-BD31-4B8C-83A1-F6EECF244321}\">" +
                    "<p14:modId xmlns:p14=\"http://schemas.microsoft.com/office/powerpoint/2010/main\" val=\"3772319657\"/>" +
                    "</p:ext>" +
                    "</p:extLst>" +
                    "</p:nvPr>" +
                    "</p:nvGraphicFramePr>";

    private final static String xlsEmbedTmpl =
            "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                    "xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<mc:Choice xmlns:v=\"urn:schemas-microsoft-com:vml\" Requires=\"v\">" +
                    "<p:oleObj spid=\"_x0000_s1028\" name=\"Arbeitsblatt\" r:id=\"@xlsId\" imgW=\"2276410\" imgH=\"1628910\" progId=\"Excel.Sheet.12\">" +
                    "<p:embed/>" +
                    "</p:oleObj>" +
                    "</mc:Choice>" +
                    "<mc:Fallback>" +
                    "<p:oleObj name=\"Arbeitsblatt\" r:id=\"@xlsId\" imgW=\"2276410\" imgH=\"1628910\" progId=\"Excel.Sheet.12\">" +
                    "<p:embed/>" +
                    "<p:pic>" +
                    "<p:nvPicPr>" +
                    "<p:cNvPr id=\"0\" name=\"\"/>" +
                    "<p:cNvPicPr/>" +
                    "<p:nvPr/>" +
                    "</p:nvPicPr>" +
                    "<p:blipFill>" +
                    "<a:blip r:embed=\"@imgId\"/>" +
                    "<a:stretch>" +
                    "<a:fillRect/>" +
                    "</a:stretch>" +
                    "</p:blipFill>" +
                    "<p:spPr>" +
                    "<a:xfrm>" +
                    "<a:off x=\"3352800\" y=\"3886200\"/>" +
                    "<a:ext cx=\"2611438\" cy=\"1506537\"/>" +
                    "</a:xfrm>" +
                    "<a:prstGeom prst=\"rect\">" +
                    "<a:avLst/>" +
                    "</a:prstGeom>" +
                    "</p:spPr>" +
                    "</p:pic>" +
                    "</p:oleObj>" +
                    "</mc:Fallback>" +
                    "</mc:AlternateContent>";
}
