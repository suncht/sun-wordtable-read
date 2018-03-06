package com.suncht.convert;

import java.io.File;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.artofsolving.jodconverter.OfficeDocumentConverter;
import org.artofsolving.jodconverter.document.DocumentFamily;
import org.artofsolving.jodconverter.document.DocumentFormat;
import org.artofsolving.jodconverter.office.OfficeManager;

/**
 * 通用文档转换，可支持docx文档
 * 支持docx文档，需要LibreOffice直接支持， 而OpenOffice不能直接支持Docx文档，需要AccessODF插件
 * LibreOffice/OpenOffice能支持哪些文档转换，那么该程序能支持哪些转换
 * @author changtan.sun
 *
 */
public class CommonDocumentConverter implements IOfficeDocumentConverter {
	protected OfficeManager officeManager;
	protected String inputFile;
	protected String outputFile;
	
	protected boolean needTempFile = false; //是否删除中间文件
	protected String tempFile;
	protected boolean needDeleteInputFile = false; //是否需要删除输入文件
	protected String extraOutputFormatToNeed;
	
	//额外的输出文档格式
	private static Map<String, String> extraOutputFormatMap = new HashMap<String, String>();
	
	static {
		//增加Docx文档格式
		extraOutputFormatMap.put("docx", "MS Word 2007 XML");
	}
	
	public CommonDocumentConverter(OfficeManager officeManager, String inputFile, String outputFile, boolean needDeleteInputFile) {
		this.officeManager = officeManager;
		this.inputFile = inputFile;
		this.outputFile = outputFile;
		this.needDeleteInputFile = needDeleteInputFile;
	}
	
	public void before() {
		tempFile = null;
		needTempFile = false;
		extraOutputFormatToNeed = null;
		
		String sufix = FileUtils.getFileSufix(outputFile);
		this.judgeFormat(sufix);
	}

	@Override
	public void convert() {
		OfficeDocumentConverter converter = new OfficeDocumentConverter(officeManager);  
		
		System.out.println("转换前处理...");  
		before();
		
        //开始转换 
		System.out.println("转换开始执行，["+inputFile+"]转换为["+outputFile+"]...");  
		if(StringUtils.isNotBlank(extraOutputFormatToNeed)) {
			DocumentFormat extraFormat = converter.getFormatRegistry().getFormatByExtension(extraOutputFormatToNeed);
			extraFormat.setStoreProperties(DocumentFamily.TEXT, Collections.singletonMap("FilterName", extraOutputFormatMap.get(extraOutputFormatToNeed)));
		
            if(needTempFile) {
				converter.convert(new File(tempFile),new File(outputFile), extraFormat);  
			} else {
				converter.convert(new File(inputFile),new File(outputFile), extraFormat);  
			}
            
		} else {
			if(needTempFile) {
				converter.convert(new File(tempFile),new File(outputFile));  
			} else {
				converter.convert(new File(inputFile),new File(outputFile));  
			}
		}
		
		System.out.println("转换后处理...");  
		after();
		
		System.out.println("转换完成");  
	}

	public void after() {
		if(needTempFile) {
			FileUtils.deleteFile(tempFile);
		}
		
		if(needDeleteInputFile) {
			FileUtils.deleteFile(inputFile);
		}
	}
	
	private void judgeFormat(String sufix) {
		if(extraOutputFormatMap.containsKey(sufix)) {
			extraOutputFormatToNeed = sufix;
		}
	}
	

}
