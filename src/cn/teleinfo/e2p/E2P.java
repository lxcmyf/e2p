package cn.teleinfo.e2p;

import java.io.File;
import java.io.FilenameFilter;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class E2P {
	
	static final int wdDoNotSaveChanges = 0;// ����������ĸ���
	static final int wdFormatPDF = 17;// wordתPDF ��ʽ 
	

	/**
	 * Excelת����PDF
	 * @param inputFile
	 * @param pdfFile
	 * @return
	 */
	public static boolean Ex2PDF(String inputFile, String pdfFile) {
		try {
			long start = System.currentTimeMillis();  
			ComThread.InitSTA(true);
			ActiveXComponent ax = new ActiveXComponent("KET.Application");
			System.out.println("��ʼת��ExcelΪPDF...");
			ax.setProperty("Visible", false);
			ax.setProperty("AutomationSecurity", new Variant(3)); // ���ú�
			Dispatch excels = ax.getProperty("Workbooks").toDispatch();
			System.out.println("���ĵ���" + excels);  
			Dispatch excel = Dispatch
					.invoke(excels, "Open", Dispatch.Method,
							new Object[] { inputFile, new Variant(false), new Variant(false) }, new int[9])
							.toDispatch();
			// ת����ʽ
			Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method, new Object[] { new Variant(0), // PDF��ʽ=0
					pdfFile, new Variant(0) // 0=��׼ (���ɵ�PDFͼƬ�����ģ��) 1=��С�ļ�
			}, new int[1]);
			System.out.println("ת���ĵ���PDF��" + pdfFile);  
			Dispatch.call(excel, "Close", new Variant(false));
			long end = System.currentTimeMillis();  
			System.out.println("ת����ɣ���ʱ��" + (end - start) + "ms");  
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
