package templatetopdf;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

import java.io.File;
import java.net.ConnectException;
import java.util.HashMap;
import java.util.Map;

/**
 * <p>类描述：将word转为pdf。</p>
 * @author贾亮
 * @version v1.0.0.1。
 * @since JDK1.8。
 * <p>创建日期：2017年2月28日 下午8:54:08。</p>
 */
public class WordToPDF {
	/*
	 * 初始化启用OpenOffice服务线程，如果是在Windows环境下运行，该静态块一般放在主函数中，
	 * 项目启动运行一次就可以了，但是在Linux环境下则不行，要手动启动OpenOffice
	 * 服务，命令如下：
	 * nohup  /opt/openoffice4/program/soffice -headless -accept="socket,host=127.0.0.1,port=8100;urp;" -nofirststartwizard &
	 */
//	static {
//		String command = "C:/Program Files (x86)/OpenOffice 4/program/soffice.exe -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";
//        try {
//			Process p = Runtime.getRuntime().exec(command);
//		} catch (IOException e) {
//			e.printStackTrace();
//		}
//	}
	
	/**
	 * <p>功能描述：调用OpenOffice将word转为PDF。</p>
	 * <p>贾亮</p>
	 * @param wordFile	原word文件
	 * @param pdfFile	需要产生的pdf文件
	 * @since JDK1.8
	 * <p>创建日期：2017/7/18 10:57。</p>
	 * <p>更新日期:[日期YYYY-MM-DD][更改人姓名][变更描述]。</p>
	 */
	public static Map<String, Object> convertPDF(File wordFile, File pdfFile) throws ConnectException {
		Map<String, Object> result = new HashMap<String, Object>();
		if (!wordFile.exists()) {
			result.put("success", -1);
			result.put("msg", "word文件丢失");
			return result;
		}
		if (!pdfFile.exists()) {
			result.put("success", -1);
			result.put("msg", "目标pdf文件丢失");
			return result;
		}
		OpenOfficeConnection connection = new SocketOpenOfficeConnection(8100);
		connection.connect();
		System.out.println("创建链接");
		DocumentConverter converter = new OpenOfficeDocumentConverter(connection);
		converter.convert(wordFile, pdfFile);
		//删除产生的临时word文件
		if (wordFile.exists()) {
			wordFile.delete();
			System.out.println("临时word文件删除成功");
		}
		if (connection != null) {
			connection.disconnect();
			System.out.println("销毁链接");
		}
		result.put("success", 1);
		result.put("msg", "转换成功");
		return result;
	}

}
