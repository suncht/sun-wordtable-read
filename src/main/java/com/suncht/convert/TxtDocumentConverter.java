package com.suncht.convert;

import java.io.File;
import java.io.FileNotFoundException;

import org.artofsolving.jodconverter.office.OfficeManager;

/**
 * 支持Txt类型文档转换
 * txt需要先转为odt类型文档，才能进行下一步转换
 * @author changtan.sun
 *
 */
public class TxtDocumentConverter extends CommonDocumentConverter {

	public TxtDocumentConverter(OfficeManager officeManager, String inputFile, String outputFile, boolean needDeleteInputFile) {
		super(officeManager, inputFile, outputFile, needDeleteInputFile);
	}

	@Override
	public void before() {
		super.before();
		
		if(inputFile.endsWith(".txt")){ //如果是Txt文件，需要转换为odt文件
			needTempFile = true;
			tempFile = FileUtils.getFilePrefix(inputFile)+".odt";
            if(new File(tempFile).exists()){
                System.out.println("odt文件已存在！");
                inputFile = tempFile;
            }else{
                try {
                    FileUtils.copyFile(inputFile, tempFile);
                    inputFile = tempFile;
                } catch (FileNotFoundException e) {
                    System.out.println("文档不存在！");
                    e.printStackTrace();
                }
            }
        }
	}
}
