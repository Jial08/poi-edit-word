package templatetopdf;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.renderer.category.LineAndShapeRenderer;
import org.jfree.data.category.CategoryDataset;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.ui.RectangleEdge;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

/**
 * 用JFreeChart生成图片
 *
 * 注意：如果项目在linux部署，可能会出现图片上的中文都显示口的情况，具体解决思如如下：
 * http://blog.csdn.net/loveany121/article/details/7944077
 */
public class LineChart {
	// 临时文件存放路径
	public static String tempfilepath = "tempFile";

	public static DefaultCategoryDataset createDate() {
		DefaultCategoryDataset mDataset = new DefaultCategoryDataset();
		mDataset.addValue(1,"y1","x1");
		mDataset.addValue(2,"y1","x2");
		mDataset.addValue(3,"y1","x3");
		mDataset.addValue(4,"y1","x4");
		mDataset.addValue(5,"y1","x5");
		mDataset.addValue(2,"y2","x1");
		mDataset.addValue(3,"y2","x2");
		mDataset.addValue(4,"y2","x3");
		mDataset.addValue(5,"y2","x4");
		mDataset.addValue(6,"y2","x5");
		return mDataset;
	}

	/*
	 * 整个大的框架属于JFreeChart
	 * 
	 * 坐标轴里的属于 Plot 其常用子类有：CategoryPlot, MultiplePiePlot, PiePlot , XYPlot
	 */
	/**
	 * <p>功能描述：用JFreeChart生成图片。</p>
	 * <p>Jial08</p>
	 * @param picHeader	图片名
	 * @param mDataset	数据
	 * @param x			x轴名
	 * @param y			y轴名
	 * @since JDK1.8
	 * <p>创建日期：2017/7/18 11:11。</p>
	 * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
	 */
	public static Map<String, Object> createChart(String picHeader, CategoryDataset mDataset, String x, String y) throws IOException {
		/*****设置主题的样式，解决中文乱码问题*************/
		// 创建主题样式
		StandardChartTheme standardChartTheme = new StandardChartTheme("CN");
		// 设置标题字体
		standardChartTheme.setExtraLargeFont(new Font("宋体", Font.BOLD, 20));
		// 设置图例的字体
		standardChartTheme.setRegularFont(new Font("宋体", Font.PLAIN, 15));
		// 设置轴向的字体
		standardChartTheme.setLargeFont(new Font("宋体", Font.PLAIN, 15));
		// 应用主题样式
		ChartFactory.setChartTheme(standardChartTheme);
		/*****设置主题的样式，解决中文乱码问题*************/

		// 定义图表对象
		JFreeChart chart = ChartFactory.createLineChart(picHeader, // 报表题目，字符串类型
				x, // 横轴
				y, // 纵轴
				mDataset = createDate(), // 获得数据集
				PlotOrientation.VERTICAL, // 图表方向垂直
				true, // 显示图例
				false, // 不用生成工具
				false // 不用生成URL地址
		);

		// 设置图例位置
		chart.getLegend().setVisible(true);
		chart.getLegend().setPosition(RectangleEdge.RIGHT);

		// 生成图形
		CategoryPlot plot = chart.getCategoryPlot();

		// 图像属性部分
		plot.setBackgroundPaint(Color.white); // 设置绘图区背景色
		plot.setDomainGridlinesVisible(false); // 设置背景网格线是否可见
		// plot.setDomainGridlinePaint(Color.BLACK); // 设置垂直方向背景线颜色
		// plot.setRangeGridlinePaint(Color.RED); // 设置水平方向背景线颜色
		plot.setNoDataMessage("没有数据");// 没有数据时显示的文字说明。

		// 数据轴属性部分，不设置则自适应
		// NumberAxis rangeAxis = (NumberAxis) plot.getRangeAxis();
		// rangeAxis.setStandardTickUnits(NumberAxis.createIntegerTickUnits());
		// rangeAxis.setAutoRangeIncludesZero(true); // 自动生成
		// rangeAxis.setUpperMargin(0.20);
		// rangeAxis.setLabelAngle(Math.PI / 2.0);
		// rangeAxis.setAutoRange(true);

		// 数据渲染部分 主要是对折线做操作
		LineAndShapeRenderer renderer = (LineAndShapeRenderer) plot.getRenderer();
		// 数据点样式设置
		renderer.setBaseShapesVisible(true); // 数据点显示外框
		renderer.setBaseShapesFilled(true); // 数据点外框内是否填充
		// renderer.setSeriesFillPaint(0, Color.ORANGE); //
		// 第一条序列线上数据点外框内填充颜色为橘黄色
		// renderer.setSeriesFillPaint(1, Color.white); // 第二条序列线上数据点外框内填充颜色为白色
		// renderer.setUseFillPaint(true); // 如果要在数据点外框内填充自定义的颜色，这个标志位必须为真
		// 序列线样式设置
		// renderer.setSeriesPaint(0, Color.GREEN); // 设置第一条序列线为绿色
		// renderer.setSeriesPaint(1, Color.YELLOW); // 设置第二条数据线为黄色

		/*
		 * 
		 * 这里的StandardCategoryItemLabelGenerator()我想强调下：当时这个地*方被搅得头很晕，
		 * Standard**ItemLabelGenerator是通用的 因为我创建*的是CategoryPlot
		 * 所以很多设置都是Category相关
		 * 
		 * 而XYPlot 对应的则是 ： StandardXYItemLabelGenerator
		 * 
		 */

		// 对于编程人员 这种根据一种类型方法联想到其他类型相似方法的思

		// 想是必须有的吧！目前只能慢慢培养了。。

		plot.setRenderer(renderer);

		// 区域渲染部分
		// double lowpress = 4.5;
		// double uperpress = 8; // 设定正常血糖值的范围
		// IntervalMarker inter = new IntervalMarker(lowpress, uperpress);
		// inter.setLabelOffsetType(LengthAdjustmentType.EXPAND); // 范围调整——扩张
		// inter.setPaint(Color.LIGHT_GRAY);// 域顏色
		// inter.setLabelFont(new Font("SansSerif", 41, 14));
		// inter.setLabelPaint(Color.RED);
		// inter.setLabel("8"); // 设定区域说明文字
		// plot.addRangeMarker(inter, Layer.BACKGROUND); // 添加mark到图形
		// BACKGROUND使得数据折线在区域的前端

		// 创建文件输出流，为了不覆盖文件，为UUID作为文件名
		String objectid = UUID.randomUUID().toString();
		File picPath = new File(tempfilepath);
		if (!picPath.exists()) {
			picPath.mkdirs();
		}
		File picFile = new File(picPath, objectid + ".jpg");
		// 输出到哪个输出流
		ChartUtilities.saveChartAsJPEG(picFile, chart, // 统计图表对象
				900, // 宽
				400 // 高
		);
		// {width=宽度,height=高度,fileType=图片格式,is=图片输入流(InputStream)}
		InputStream is = new FileInputStream(picFile);
		Map<String, Object> picMap = new HashMap<String, Object>();
		picMap.put("width", 900);
		picMap.put("height", 400);
		picMap.put("fileType", "jpg");
		picMap.put("is", is);
		picMap.put("picFile", picFile);
		return picMap;
		// 窗口查看效果图
		// ChartFrame mChartFrame = new ChartFrame("折线图", chart);
		// mChartFrame.pack();
		// mChartFrame.setVisible(true);

	}

}
