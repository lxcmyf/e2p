package cn.teleinfo.e2p;

import java.io.File;
import java.io.FilenameFilter;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class E2P {
	
	static final int wdDoNotSaveChanges = 0;// 不保存待定的更改
	static final int wdFormatPDF = 17;// word转PDF 格式 
	

	/**
	 * Excel转化成PDF
	 * @param inputFile
	 * @param pdfFile
	 * @return
	 */
	public static boolean Ex2PDF(String inputFile, String pdfFile) {
		try {
			long start = System.currentTimeMillis();  
			ComThread.InitSTA(true);
			ActiveXComponent ax = new ActiveXComponent("KET.Application");
			System.out.println("开始转化Excel为PDF...");
			ax.setProperty("Visible", false);
			ax.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
			Dispatch excels = ax.getProperty("Workbooks").toDispatch();
			System.out.println("打开文档：" + excels);  
			Dispatch excel = Dispatch
					.invoke(excels, "Open", Dispatch.Method,
							new Object[] { inputFile, new Variant(false), new Variant(false) }, new int[9])
							.toDispatch();
			// 转换格式
			Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method, new Object[] { new Variant(0), // PDF格式=0
					pdfFile, new Variant(0) // 0=标准 (生成的PDF图片不会变模糊) 1=最小文件
			}, new int[1]);
			System.out.println("转换文档到PDF：" + pdfFile);  
			Dispatch.call(excel, "Close", new Variant(false));
			long end = System.currentTimeMillis();  
			System.out.println("转换完成，用时：" + (end - start) + "ms");  
			return true; 
		} catch (Exception e) {
			return false;
		}
	}

	
	public static void main(String[] args) {
		String[] excelNames = new File("d:/").list(new FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.endsWith(".xlsx");
            }
        });
		for (String string : excelNames) {
			Ex2PDF("d://"+string, "D://"+Math.random()+".pdf");
		}
	}


}
