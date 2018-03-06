package com.suncht.convert.demo;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.Collections;

import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

import com.suncht.convert.FileUtils;
  
public class OfficePDFConverter{  
      
    private static OfficePDFConverter converter = new OfficePDFConverter();  
    private static  OfficeManager officeManager;  
    //openOffice安装路径  
    //private static String OPEN_OFFICE_HOME = "D:\\Program Files (x86)\\OpenOffice 4\\";  
    private static String OPEN_OFFICE_HOME = "D:\\Program Files\\LibreOffice 5\\";  
    //服务端口  
    private static int OPEN_OFFICE_PORT[] = {8100};  
      
    public static OfficePDFConverter getConverter() {  
        return converter;  
    }  
      
    /** 
     *  
     * office2Pdf 方法 
     * @descript：TODO 
     * @param inputFile 文件全路径 
     * @param outputFile pdf文件全路径 
     * @return void 
     * @author lxz 
     * @return  
     */      
    public void convert2PDF(String inputFile,String outputFile) {  
          
    	if(inputFile.endsWith(".txt")){
            String odtFile = FileUtils.getFilePrefix(inputFile)+".odt";
            if(new File(odtFile).exists()){
                System.out.println("odt文件已存在！");
                inputFile = odtFile;
            }else{
                try {
                    FileUtils.copyFile(inputFile,odtFile);
                    inputFile = odtFile;
                } catch (FileNotFoundException e) {
                    System.out.println("文档不存在！");
                    e.printStackTrace();
                }
            }
        }
    	
        File pdfFile = new File(outputFile);  
        if (pdfFile.exists()) {  
            pdfFile.delete();  
        }  
        try{  
            long startTime = System.currentTimeMillis();  
            //打开服务  
            startService();          
            OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);  
            //开始转换  
            converter.convert(new File(inputFile),new File(outputFile));  
            //关闭  
            stopService();  
            System.out.println("运行结束");  
        }catch (Exception e) {  
            // TODO: handle exception  
            e.printStackTrace();  
        }  
    }  
    
    public static void doc2Docx(String inputFile,String outputFile) {
    	File pdfFile = new File(outputFile);  
        if (pdfFile.exists()) {  
            pdfFile.delete();  
        }  
        try{  
            long startTime = System.currentTimeMillis();  
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
      
    public static void stopService(){  
        if (officeManager != null) {  
            officeManager.stop();  
        }  
    }  
      
    public static void startService(){  
        DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();  
        try {
        	System.out.println("准备启动服务....");
            configuration.setOfficeHome(OPEN_OFFICE_HOME);//设置安装目录  
            configuration.setPortNumbers(OPEN_OFFICE_PORT); //设置端口  
            configuration.setTaskExecutionTimeout(1000 * 60 * 5L);  
            configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);  
            officeManager = configuration.buildOfficeManager();  
            officeManager.start();    //启动服务  
            System.out.println("office转换服务启动成功!");
        } catch (Exception ce) {  
            System.out.println("office转换服务启动失败!详细信息:" + ce);  
        }  
    }  
}  