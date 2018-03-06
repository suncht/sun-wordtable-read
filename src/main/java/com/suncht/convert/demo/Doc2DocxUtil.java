package com.suncht.convert.demo;
import java.io.File;
import java.util.Collections;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;
  
public class Doc2DocxUtil{  
      
    private static Doc2DocxUtil doc2DocxUtil = new Doc2DocxUtil();  
    private static  OfficeManager officeManager;  
    //openOffice安装路径  
    private static String OPEN_OFFICE_HOME = "D:\\Program Files\\LibreOffice 5\\";  
    //服务端口  
    private static int OPEN_OFFICE_PORT[] = {8101};  
      
    public static Doc2DocxUtil getOffice2PdfUtil() {  
        return doc2DocxUtil;  
    }  
      
    
    public static void doc2Docx(String inputFile,String outputFile) {
    	File pdfFile = new File(outputFile);  
        if (pdfFile.exists()) {  
            pdfFile.delete();  
        }  
        try{  
            //打开服务  
            startService();          
            OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);  
            DocumentFormat docx = converter.getFormatRegistry().getFormatByExtension("docx");
            docx.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "MS Word 2007 XML"));
            //开始转换  
            converter.convert(new File(inputFile),new File(outputFile), docx);  
            //关闭  
            stopService();  
            System.out.println("运行结束");  
        }catch (Exception e) {  
            // TODO: handle exception  
            e.printStackTrace();  
        }  
    }
    
    private static void transformBinaryWordDocToDocX(File in, File out) {
        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);
        DocumentFormat docx = converter.getFormatRegistry().getFormatByExtension("docx");
        docx.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "MS Word 2007 XML"));

        converter.convert(in, out, docx);
    }
    
    private static void transformBinaryWordDocToW2003Xml(File in, File out) {
        OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);;
        DocumentFormat w2003xml = new DocumentFormat("Microsoft Word 2003 XML", "xml", "text/xml");
        w2003xml.setInputFamily(DocumentFamily.TEXT);
        w2003xml.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", "MS Word 2003 XML"));
        converter.convert(in, out, w2003xml);
    }

      
    public static void stopService(){  
        if (officeManager != null) {  
            officeManager.stop();  
        }  
    }  
      
    public static void startService(){  
        DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();  
        try {  
            configuration.setOfficeHome(OPEN_OFFICE_HOME);//设置安装目录  
            configuration.setPortNumbers(OPEN_OFFICE_PORT); //设置端口  
            configuration.setTaskExecutionTimeout(1000 * 60 * 5L);  
            configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);  
            officeManager = configuration.buildOfficeManager();  
            officeManager.start();    //启动服务  
        } catch (Exception ce) {  
            System.out.println("office转换服务启动失败!详细信息:" + ce);  
        }  
    }  
}  