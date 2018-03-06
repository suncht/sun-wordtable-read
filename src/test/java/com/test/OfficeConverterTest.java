package com.test;

import org.junit.Test;

import com.suncht.convert.OfficeDocumentConvertServer;

public class OfficeConverterTest {
	private static String OPEN_OFFICE_HOME = "D:\\Program Files\\LibreOffice 5\\";
	private static int OPEN_OFFICE_PORT[] = { 8101 };
	
	@Test
	public void txt2docx() {
		String inputFile = "D:\\dic.txt";
		String outputFile = "D:\\dic.docx";

		// 服务端口
		try (OfficeDocumentConvertServer server = new OfficeDocumentConvertServer(OPEN_OFFICE_HOME, OPEN_OFFICE_PORT);) {
			server.convert(inputFile, outputFile, false);
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	@Test
	public void docx2pdf() {
		String inputFile = "D:\\故障模式分析表格样例 - 副本.docx";
		String outputFile = "D:\\故障模式分析表格样例 - 副本.pdf";

		// 服务端口
		try (OfficeDocumentConvertServer server = new OfficeDocumentConvertServer(OPEN_OFFICE_HOME, OPEN_OFFICE_PORT);) {
			server.convert(inputFile, outputFile, false);
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
}
