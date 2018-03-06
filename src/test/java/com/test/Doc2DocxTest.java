package com.test;

import com.suncht.convert.OfficeDocumentConvertServer;

public class Doc2DocxTest {

	public static void main(String[] args) throws Exception {
		String inputFile = "D:\\FMEA信息导入-客户实例.doc";
		String outputFile = "D:\\FMEA信息导入-客户实例.docx";
		//Doc2DocxUtil.doc2Docx(outputFile, inputFile);

//		Thread.sleep(2000);
		String pdfFile = "D:\\FMEA信息导入-客户实例.pdf";
//		OfficePDFConverter.getConverter().convert2PDF(outputFile, pdfFile);

		String OPEN_OFFICE_HOME = "D:\\Program Files\\LibreOffice 5\\";
		// 服务端口
		int OPEN_OFFICE_PORT[] = { 8101 };
		try (OfficeDocumentConvertServer server = new OfficeDocumentConvertServer(OPEN_OFFICE_HOME, OPEN_OFFICE_PORT);) {
			server.convert(inputFile, outputFile, false);
			server.convert(outputFile, pdfFile, true);
		}

	}

}
