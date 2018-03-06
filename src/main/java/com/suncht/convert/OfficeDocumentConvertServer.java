package com.suncht.convert;

import java.io.Closeable;
import java.io.IOException;

import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration;
import org.artofsolving.jodconverter.office.OfficeManager;

/**
 * 文档转换服务
 * @author changtan.sun
 *
 */
public class OfficeDocumentConvertServer implements Closeable {
	private OfficeManager officeManager;

	public OfficeDocumentConvertServer(String officeHome, int... ports) {
		this.startService(officeHome, ports);
	}
	
	private void startService(String officeHome, int... ports) {
		DefaultOfficeManagerConfiguration configuration = new DefaultOfficeManagerConfiguration();
		try {
			System.out.println("准备启动office服务....");
			configuration.setOfficeHome(officeHome);// 设置安装目录
			configuration.setPortNumbers(ports); // 设置端口
			configuration.setTaskExecutionTimeout(1000 * 60 * 5L);
			configuration.setTaskQueueTimeout(1000 * 60 * 60 * 24L);
			officeManager = configuration.buildOfficeManager();
			officeManager.start(); // 启动服务
			System.out.println("office转换服务启动成功!");
		} catch (Exception ce) {
			System.out.println("office转换服务启动失败!详细信息:" + ce);
		}
	}

	public void convert(String inputFile, String outputFile, boolean needDeleteInputFile) {
		IOfficeDocumentConverter converter = null;
		if(inputFile.endsWith(".txt")) {
			converter = new TxtDocumentConverter(officeManager, inputFile, outputFile, needDeleteInputFile);
		} else {
			converter = new CommonDocumentConverter(officeManager, inputFile, outputFile, needDeleteInputFile);
		}
		
		converter.convert();
	}
	
	public void convert(String inputFile, String outputFile) {
		this.convert(inputFile, outputFile, false);
	}

	@Override
	public void close() throws IOException {
		if (officeManager != null) {
			officeManager.stop();
			System.out.println("关闭office服务");  
		}
	}
}
