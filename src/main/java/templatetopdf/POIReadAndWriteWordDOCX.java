package templatetopdf;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.*;
import java.util.*;
import java.util.Map.Entry;

/**
 * <p>类描述：使用POI，读取word2007及以上版本，并实现修改文本内容，替换表格中的文本内容，在指定位置插入图片，并写回到word中。</p>
 *
 * @version v1.0.0.1。
 * @authorJial08
 * @since JDK1.8。
 * <p>创建日期：2017年4月11日 上午11:23:30。</p>
 */
public class POIReadAndWriteWordDOCX {
    // 临时文件存放路径
    public static String tempfilepath = "tempFile";

    /**
     * <p>功能描述：根据图片类型，取得对应的图片类型代码。</p>
     * <p>Jial08 </p>
     *
     * @param picType 图片格式
     * @return
     * @since JDK1.8。
     * <p>创建日期:2017年4月10日 上午11:41:30。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    private static int getPictureType(String picType) {
        int res = CustomXWPFDocument.PICTURE_TYPE_PICT;
        if (picType != null) {
            if (picType.equalsIgnoreCase("png")) {
                res = CustomXWPFDocument.PICTURE_TYPE_PNG;
            } else if (picType.equalsIgnoreCase("dib")) {
                res = CustomXWPFDocument.PICTURE_TYPE_DIB;
            } else if (picType.equalsIgnoreCase("emf")) {
                res = CustomXWPFDocument.PICTURE_TYPE_EMF;
            } else if (picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")) {
                res = CustomXWPFDocument.PICTURE_TYPE_JPEG;
            } else if (picType.equalsIgnoreCase("wmf")) {
                res = CustomXWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }

    /**
     * <p>功能描述：读取并操作word2007及以上word中的内容。</p>
     * <p>Jial08 </p>
     *
     * @param is           模版文件流
     * @param param        需要替换的文本和图片数据，其中图片信息用Map保存，格式如：{width=宽度,height=高度,fileType=图片格式,is=图片输入流(InputStream)}
     * @param wordName     需要导出的word名，需要加后缀；如果不需要导出word则写为null
     * @param allExcelData 插入表格中的数据
     * @param picMap       如果某一个位置需要插入多张图片，则先扩展该位置的变量数
     * @return result
     * @throws IOException
     * @throws InvalidFormatException
     * @since JDK1.8。
     * <p>创建日期:2017年4月10日 下午6:22:31。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    public static Map<String, Object> readwriteWord(InputStream is, Map<String, Object> param, String wordName, List<List> allExcelData, Map<String, Integer> picMap) throws IOException, InvalidFormatException {
        Map<String, Object> result = new HashMap<String, Object>();
        /**判断源文件是否存在*/
        if (is == null) {
            result.put("success", -1);
            result.put("msg", "模版文件不存在！");
            return result;
        }
        CustomXWPFDocument document;
        // 打开word2007的文件
        document = new CustomXWPFDocument(is);

        // 指定位置插入表格
        insertTab("Excel", document);

        // 获取到所有的段落（除表格外的都是段落）
        List<XWPFParagraph> listParagraphs = document.getParagraphs();
        // 扩展插入图片位置的变量
        if (picMap != null) {
            for (Entry<String, Integer> entry : picMap.entrySet()) {
                String key = entry.getKey();
                Integer value = entry.getValue();
                extendPic(document, listParagraphs, key, value);
            }
        }
        // 替换word2007的纯文本内容，即除去表格以外的段落
        result = processParagraphs(listParagraphs, param, document);
        int success = (Integer) result.get("success");
        if (success == -1) {
            return result;
        }

        // 取得所有表格，替换表格中的文字和图片
        Iterator<XWPFTable> it = document.getTablesIterator();
        while (it.hasNext()) {// 循环操作表格
            XWPFTable table = it.next();
            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {// 取得表格的行
                List<XWPFTableCell> cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {// 取得单元格，并将单元格作为段落处理
                    List<XWPFParagraph> listParagraphsTable = cell.getParagraphs();
                    result = processParagraphs(listParagraphsTable, param, document);
                    success = (Integer) result.get("success");
                    if (success == -1) {
                        return result;
                    }
                }
            }
        }

        // 向表格中插入数据
        List<XWPFTable> tables = document.getTables();
        insertTableData(tables, allExcelData);

        // 导出word
        String reportId = UUID.randomUUID().toString();
        File report = new File(tempfilepath, reportId + ".docx");
        OutputStream os = new FileOutputStream(report);
        document.write(os);
        if (os != null) {
            os.close();
        }
        result.put("success", 1);
        result.put("tempWord", report.getPath());
        return result;
    }

    /**
     * <p>功能描述：扩展需要插入图片位置的变量，即增加XWPFRun的数量。</p>
     * <p>Jial08 </p>
     *
     * @param listParagraphs
     * @param str
     * @param num
     * @since JDK1.8。
     * <p>创建日期:2017年5月31日 下午2:58:49。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    private static void extendPic(CustomXWPFDocument document, List<XWPFParagraph> listParagraphs, String str, int num) {
        String newStr = str.substring(0, str.length() - 1);
        if (listParagraphs != null && listParagraphs.size() > 0) {
            for (int i = 0; i < listParagraphs.size(); i++) {
                List<XWPFRun> runs = listParagraphs.get(i).getRuns();
                for (int k = 0; k < runs.size(); k++) {
                    XWPFRun run = runs.get(k);
                    String text = run.getText(0);
                    if (("${" + str + "}").equals(text)) {
                        if (num <= 1) {
                            continue;
                        }
                        for (int j = 0; j < num - 1; j++) {
                            XWPFParagraph paragraph = listParagraphs.get(i);
                            XWPFRun newRun = paragraph.createRun();
                            newRun.setText("${" + newStr + (j + 2) + "}");
                            /*
                             * 如果要使图片都换行显示，则需要为新增的run换行，如果需要在一行显示,
							 * 则下面的代码注释掉即可。
							 */
                            if (j < num - 1) {
                                newRun.addBreak();
//								newRun.addCarriageReturn();//也可以换行
                            }
                        }
                    }
                }
            }

        }
    }

    /**
     * <p>功能描述：替换段落中的文字和图片。</p>
     * <p>Jial08 </p>
     *
     * @param listParagraphs 段落集合
     * @param param          模版中需要插入的文字和图片的Map集合
     * @param document       word文件对象
     * @return
     * @throws InvalidFormatException
     * @since JDK1.8。
     * <p>创建日期:2017年4月10日 下午1:27:46。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    private static Map<String, Object> processParagraphs(List<XWPFParagraph> listParagraphs, Map<String, Object> param, CustomXWPFDocument document) throws InvalidFormatException {
        Map<String, Object> result = new HashMap<String, Object>();
        if (listParagraphs != null && listParagraphs.size() > 0) {
            for (XWPFParagraph paragraph : listParagraphs) {
                // 获取段落文本
                String paragraphText = paragraph.getParagraphText();
                // 如果段落中不包含${}变量标识符则跳过
                if (!paragraphText.contains("${")) {
                    continue;
                }
                // XWPFRun对象定义了一个带有公共属性集的文本区域
                List<XWPFRun> runs = paragraph.getRuns();
                for (int i = 0; i < runs.size(); i++) {
                    XWPFRun run = runs.get(i);
                    String text = run.getText(0);
//					System.out.println(text);
                    if (text != null && !"".equals(text)) {
                        for (Entry<String, Object> entry : param.entrySet()) {
                            String key = entry.getKey();
                            if (text.equals("${" + key + "}")) {
                                Object value = entry.getValue();
                                if (!(value instanceof Map)) {// 文本替换
                                    /*
									 * run.setText(value); 在指定位置追加文本
									 * run.setText(value, pos); 替换指定位置的文本
									 */
                                    if (value == null) {
                                        run.setText("", 0);
                                    } else {
                                        run.setText(value.toString(), 0);
                                    }
                                } else if (value instanceof Map) {// 图片替换
                                    Map picture = (Map) value;
                                    int width = Integer.parseInt(picture.get("width").toString());
                                    int height = Integer.parseInt(picture.get("height").toString());
                                    String fileType = picture.get("fileType").toString();
                                    int picType = getPictureType(fileType);
                                    InputStream is = null;
                                    is = (InputStream) picture.get("is");
                                    if (is == null) {
                                        result.put("success", -1);
                                        result.put("msg", "图片文件不存在！");
                                        return result;
                                    }
                                    document.addPictureData(is, picType);
                                    document.createPicture(document.getAllPictures().size() - 1, width, height, run);
                                }
                            }
                        }
                    }
                }
            }
        }
        result.put("success", 1);
        result.put("msg", "替换段落中的文字和图片成功");
        return result;
    }

    /**
     * <p>功能描述：向word表格中插入数据。</p>
     * <p>Jial08 </p>
     *
     * @param tables       word中表格集合
     * @param allExcelData 需要向表格中插入的数据集合
     * @since JDK1.8。
     * <p>创建日期:2017年6月1日 上午11:30:39。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    private static void insertTableData(List<XWPFTable> tables, List<List> allExcelData) {
        for (int i = 0; i < tables.size(); i++) {
            XWPFTable table = tables.get(i);
            List listData = allExcelData.get(i);
            // 如果没有数据则跳过该表格直接操作下一个表格
            if (listData == null || listData.size() < 1) {
                continue;
            }
            List<XWPFTableRow> rows = table.getRows();
            // 默认有一行空行可以插入数据，如果需要插入的数据大于1条，则先扩展表格
            if (listData.size() > 1) {
                for (int j = 0; j < listData.size() - 1; j++) {
                    table.createRow();
					/*
					 * 追加表格经试验只有上面的方法可行，网上说下面的方法追加的行会保持已创建
					 * 行的格式，但经测试无效，发现jar包中该方法未实现，待进一步确认。
					 */
                    // table.addNewRowBetween(rows.size() - 2, rows.size() - 1);
                }
            }
            if (listData.size() > 0) {
                // 表头所占行数
                int num = 0;
                // 列数
                int cellNum = 0;
                for (int k = 0; k < rows.size(); k++) {
                    XWPFTableRow row = rows.get(k);
					/*
					 * 跳过表头，从空行开始插入数据，如果存在第一列表格数小于后几列，
					 * 那么即使仍在表头列也会存在row.getCell(0).getText()不存的情况，
					 * 所以为了保险起见，整一行为空才为空行，即填写数据的第一行。
					 */
                    String text = "";
                    List<XWPFTableCell> cells = row.getTableCells();
                    for (XWPFTableCell cell : cells) {
                        text += cell.getText();
                    }
                    if (text.length() > 0) {
                        num++;
                        continue;
                    }
                    List excelData = (List) listData.get(k - num);
					/*
					 * 创建新行默认按照表格的第一行追加，但第一行多为表头，存在合并单元格情况，
					 * 那么就按照上一行的列数为标准，如果新增行的列数少于上一行，则追加剩余列数。 for循环中为什么不用x <
					 * cellNum - cells.size()而是多一个参数temp是因为
					 * 在新增列后cells.size()大小会跟着变化。
					 */
                    int temp = cells.size();
                    if (cells.size() < cellNum) {
                        for (int x = 0; x < cellNum - temp; x++) {
                            row.addNewTableCell();
                        }
                    }
                    cellNum = cells.size();
                    for (int l = 0; l < cellNum; l++) {
                        XWPFTableCell cell = cells.get(l);
                        // 设置表格内容水平垂直居中
                        setCellText(cell);
                        Object obj = excelData.get(l);
                        String value = obj == null ? "" : obj.toString();
                        cell.setText(value);
                    }
                }

            }
        }
    }

    /**
     * <p>功能描述：设置word中表格内容水平和垂直居中。</p>
     * <p>Jial08 </p>
     *
     * @param cell
     * @since JDK1.8。
     * <p>创建日期:2017年5月27日 上午9:14:55。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    private static void setCellText(XWPFTableCell cell) {
        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        CTShd ctshd = ctPr.addNewShd();
        // 垂直居中
        ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
		/*
		 * 水平居中 下面的方法需要引入ooxml-schemas-1.X.jar包而不是poi-ooxml-schemas-XX.jar
		 * 下载地址：https://repo1.maven.org/maven2/org/apache/poi/ooxml-schemas/1.3/
		 */
//		List<CTP> listCTP = cttc.getPList();
//		CTP ctp = listCTP.get(0);
//		CTPPr cr = ctp.addNewPPr();
//		CTJc cc = cr.addNewJc();
//		cc.setVal(STJc.CENTER);
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);

    }

    /**
     * <p>功能描述：根据检测方法删除多余表格。</p>
     * <p>Jial08 </p>
     *
     * @param is
     * @param detectMethod
     * @return
     * @throws IOException
     * @since JDK1.8。
     * <p>创建日期:2017年6月12日 上午11:07:10。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    public static String changeTable(InputStream is, int detectMethod) throws IOException {
        XWPFDocument document = new XWPFDocument(is);
        List<XWPFParagraph> listp = document.getParagraphs();
        for (int i = 0; i < listp.size(); i++) {
            XWPFParagraph paragraph = listp.get(i);
            String text = paragraph.getParagraphText();
            // 删除多余表格
            if (!("M" + String.valueOf(detectMethod)).equals(text) && text.indexOf("M") == 0 && text.length() <= 3) {
                int pos = document.getPosOfParagraph(paragraph);
                document.removeBodyElement(pos);
                i--;
                document.removeBodyElement(pos);
                i--;
            }
            // 删除表格标识符
            if (("M" + String.valueOf(detectMethod)).equals(text)) {
                int pos = document.getPosOfParagraph(paragraph);
                document.removeBodyElement(pos);
                i--;
            }
            // 删除多余段落
            if (detectMethod == 2 || detectMethod == 3 || detectMethod == 6 || detectMethod == 7) {
                if (text.indexOf("（4）系统依据判据") != -1) {
                    int pos = document.getPosOfParagraph(paragraph);
                    document.removeBodyElement(pos);
                    i--;
                    document.removeBodyElement(pos);
                    i--;
                } else if (text.indexOf("（5）其他发现") != -1) {
                    List<XWPFRun> list = paragraph.getRuns();
                    for (XWPFRun run : list) {
                        String runText = run.getText(0);
                        if ("（5）".equals(runText)) {
                            run.setText("（4）", 0);
                        }
                    }
                }
            } else if (detectMethod == 4 || detectMethod == 5 || detectMethod == 8 || detectMethod == 9 || detectMethod == 10) {
                if (text.indexOf("（3）系统依据判据") != -1 || text.indexOf("（4）系统依据判据") != -1) {
                    int pos = document.getPosOfParagraph(paragraph);
                    document.removeBodyElement(pos);
                    i--;
                    document.removeBodyElement(pos);
                    i--;
                } else if (text.indexOf("（5）其他发现") != -1) {
                    List<XWPFRun> list = paragraph.getRuns();
                    for (XWPFRun run : list) {
                        String runText = run.getText(0);
                        if ("（5）".equals(runText)) {
                            run.setText("（3）", 0);
                        }
                    }
                }
            }
        }

        deleteParagraph(document, listp, detectMethod);

        String fileId = UUID.randomUUID().toString();
        File file = new File(tempfilepath);
        if (!file.exists()) {
            file.mkdirs();
        }
        String filePath = tempfilepath + File.separator + fileId + ".docx";
        OutputStream os = new FileOutputStream(filePath);
        document.write(os);
        if (is != null) {
            is.close();
        }
        if (os != null) {
            os.close();
        }
        return filePath;
    }

    /**
     * <p>功能描述：删除“2.阴极保护环境检测方法”中多余的段落，段落以“m1.....m1”的方式标识，
     * 删除两个标识符之间的一切内容，包括这两个标识符。</p>
     * <p>Jial08 </p>
     *
     * @param document
     * @param listp
     * @param detectMethod 检测方法
     * @since JDK1.8。
     * <p>创建日期:2017年6月14日 下午2:01:00。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    public static void deleteParagraph(XWPFDocument document, List<XWPFParagraph> listp, int detectMethod) {
        boolean delete = false;
        boolean first = false;
        for (int i = 0; i < listp.size(); i++) {
            XWPFParagraph paragraph = listp.get(i);
            String text = paragraph.getParagraphText();
            // 删除指定标识符以外的标识符
            if (!("m" + String.valueOf(detectMethod)).equals(text) && text.indexOf("m") == 0 && text.length() <= 3) {
                delete = !delete;
                first = !first;
                int pos = document.getPosOfParagraph(paragraph);
                document.removeBodyElement(pos);
                i--;
                continue;
            }
            // 删除指定标识符
            if (("m" + String.valueOf(detectMethod)).equals(text)) {
                int pos = document.getPosOfParagraph(paragraph);
                document.removeBodyElement(pos);
                i--;
            }
            // 删除标识符之间的段落
            if (delete && first) {
                int pos = document.getPosOfParagraph(paragraph);
                document.removeBodyElement(pos);
                i--;
            }
        }
    }

    /**
     * <p>功能描述：向key所在位置插入表格。</p>
     * <p>Jial08</p>
     *
     * @param key
     * @param doc2
     * @since JDK1.8
     * <p>创建日期：2017/7/18 16:05。</p>
     * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
     */
    public static void insertTab(String key, XWPFDocument doc2) {
        List<XWPFParagraph> paragraphList = doc2.getParagraphs();
        if (paragraphList != null && paragraphList.size() > 0) {

            for (XWPFParagraph paragraph : paragraphList) {
                List<XWPFRun> runs = paragraph.getRuns();

                for (XWPFRun run : runs) {
                    String text = run.getText(0);
                    if (text != null) {
                        if (text.indexOf(key) >= 0) {
                            // 将需要替换的字符串清除
                            run.setText("", 0);

                            XmlCursor cursor = paragraph.getCTP().newCursor();

                            XWPFTable tableOne = doc2.insertNewTbl(cursor);// ---这个是关键

                            // XWPFTable tableOne =
                            // paragraph.getDocument().createTable();

                            XWPFTableRow tableOneRowOne = tableOne.getRow(0);
                            tableOneRowOne.getCell(0).setText("一行一列");
                            XWPFTableCell cell12 = tableOneRowOne.createCell();
                            cell12.setText("一行二列");
                            // tableOneRowOne.addNewTableCell().setText("第1行第2列");
                            // tableOneRowOne.addNewTableCell().setText("第1行第3列");
                            // tableOneRowOne.addNewTableCell().setText("第1行第4列");

                            XWPFTableRow tableOneRowTwo = tableOne.createRow();
                            tableOneRowTwo.getCell(0).setText("第二行第一列");
                            tableOneRowTwo.getCell(1).setText("第二行第二列");
                            // tableOneRowTwo.getCell(2).setText("第2行第3列");

                            XWPFTableRow tableOneRow3 = tableOne.createRow();
                            //---顺序增加行后，忽略第1、2单元格，直接插入3、4
                            tableOneRow3.addNewTableCell().setText("第三行第3列");
                            tableOneRow3.addNewTableCell().setText("第三行第4列");

                        }
                    }
                }

            }
        }
    }

}
