package templatetopdf;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

import java.io.IOException;
import java.io.InputStream;

public class CustomXWPFDocument extends XWPFDocument {
    public CustomXWPFDocument(InputStream in) throws IOException {
        super(in);
    }
    
    public CustomXWPFDocument() {
        super();
    }
    
    public CustomXWPFDocument(OPCPackage pkg) throws IOException {
        super(pkg);
    }
    
    /**
     * <p>功能描述：向段落和word表格中插入图片，目前存在一个问题，插入重复图片会使插入的图片错乱。</p>
     * <p>贾亮 </p>	
     * @param id		CustomXWPFDocument中图片位置
     * @param width		图片宽度
     * @param height	图片高度
     * @param run		带有公共属性集的文本区域
     * @since JDK1.8。
     * <p>创建日期:2017年4月10日 下午1:11:50。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    public void createPicture(int id, int width, int height,XWPFRun run) {
        final int EMU = 9525;
        width *= EMU;
        height *= EMU;
        String blipId = getAllPictures().get(id).getPackageRelationship()
                .getId();    
    
        /*
         * 如果直接在run后面画图会出现变量+图的情况，所以应先清空run的内容
         */
        run.setText("", 0);
        
        CTInline inline = run.getCTR()
                .addNewDrawing().addNewInline();    
    
        String picXml = ""    
                + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"    
                + "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"    
                + "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">"    
                + "         <pic:nvPicPr>" + "            <pic:cNvPr id=\""    
                + id    
                + "\" name=\"Generated\"/>"    
                + "            <pic:cNvPicPr/>"    
                + "         </pic:nvPicPr>"    
                + "         <pic:blipFill>"    
                + "            <a:blip r:embed=\""    
                + blipId    
                + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"    
                + "            <a:stretch>"    
                + "               <a:fillRect/>"    
                + "            </a:stretch>"    
                + "         </pic:blipFill>"    
                + "         <pic:spPr>"    
                + "            <a:xfrm>"    
                + "               <a:off x=\"0\" y=\"0\"/>"    
                + "               <a:ext cx=\""    
                + width    
                + "\" cy=\""    
                + height    
                + "\"/>"    
                + "            </a:xfrm>"    
                + "            <a:prstGeom prst=\"rect\">"    
                + "               <a:avLst/>"    
                + "            </a:prstGeom>"    
                + "         </pic:spPr>"    
                + "      </pic:pic>"    
                + "   </a:graphicData>" + "</a:graphic>";    
    
        inline.addNewGraphic().addNewGraphicData();
        XmlToken xmlToken = null;
        try {    
            xmlToken = XmlToken.Factory.parse(picXml);
        } catch (XmlException xe) {    
            xe.printStackTrace();    
        }    
        inline.set(xmlToken);    
    
        inline.setDistT(0);    
        inline.setDistB(0);    
        inline.setDistL(0);    
        inline.setDistR(0);    
    
        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);    
        extent.setCy(height);    
    
        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);    
        docPr.setName("图片" + id);    
        docPr.setDescr("descr");    
    }    
}