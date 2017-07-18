import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import templatetopdf.LineChart;
import templatetopdf.POIReadAndWriteWordDOCX;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * <p>类描述：。</p>
 *
 * @author 贾亮
 * @version v1.0.0.1。
 * @since JDK1.8。
 * <p>创建日期：2017/7/18 10:49。</p>
 */
public class test {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        InputStream is = new FileInputStream("src\\main\\resources\\templates\\inspectionRecord.docx");
        // 删除多余段落和表格并生成临时文件
        String wordPath = POIReadAndWriteWordDOCX.changeTable(is, 1);
        // 用JFreeChart生成图片
        DefaultCategoryDataset mDataset = LineChart.createDate();
        Map<String, Object> picMap = LineChart.createChart("测试用曲线图", mDataset, "x坐标", "y坐标");
        // 获取图片路径
        // 将需要替换的文字和图片放入map集合中
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("a", "哈哈");
        map.put("b", 10);
        map.put("c", "1、2和3");
        map.put("graph1", picMap);
        // 需要插入表格中的数据
        // 第一个表格中插入数据，剩余表格不动
        List<List> allExcelData = new ArrayList<List>();
        List list1 = new ArrayList();
        for (int i = 0; i < 5; i++) {
            List list = new ArrayList();
            list.add(1);
            list.add(2);
            list.add(3);
            list.add(4);
            list.add(5);
            list.add(6);
            list.add(7);
            list.add(8);
            list.add(9);
            list1.add(list);
        }
        allExcelData.add(list1);
        allExcelData.add(null);
        allExcelData.add(null);
        allExcelData.add(null);
        // 替换段落中的文字和图片，填充表格数据
        File file = new File(wordPath);
        is = new FileInputStream(file);
        POIReadAndWriteWordDOCX.readwriteWord(is, map, "测试.docx", allExcelData, null);
        // 关闭折线图的流并删除临时文件
        if (picMap.get("is") != null) {
            InputStream picIs1 = (InputStream) picMap.get("is");
            if (picIs1 != null) {
                picIs1.close();
            }
        }
        if (picMap.get("picFile") != null) {
            File picFile = (File) picMap.get("picFile");
            if (picFile.exists()) {
                picFile.delete();
            }
        }
        // 删除临时word文件
        if (file.exists()) {
            file.delete();
        }
    }
}
