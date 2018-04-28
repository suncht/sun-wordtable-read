package com.test;

import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.dev.POIFSViewEngine;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;

public class WordEmbedsTest {
	@Test
	public void listAllEmbeds() {
		try (InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/嵌套附件01.docx");) {
			XWPFDocument document = new XWPFDocument(inputStream);
			listEmbeds(document);
			//listEmbeds2(document);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void listEmbeds(XWPFDocument doc) throws OpenXML4JException {
		List<PackagePart> embeddedDocs = doc.getAllEmbedds();
		if (embeddedDocs != null && !embeddedDocs.isEmpty()) {
			Iterator<PackagePart> pIter = embeddedDocs.iterator();
			while (pIter.hasNext()) {
				PackagePart pPart = pIter.next();
				System.out.print(pPart.getPartName() + ", ");

				System.out.print(pPart.getContentType() + ", ");
				System.out.println();
			}
		}
	}

	private static void listEmbeds2(XWPFDocument doc) throws Exception {
		for (final PackagePart pPart : doc.getAllEmbedds()) {
			final String contentType = pPart.getContentType();
			System.out.println(contentType + "\n");
			if (contentType.equals("application/vnd.ms-excel")) {
				final HSSFWorkbook embeddedWorkbook = new HSSFWorkbook(pPart.getInputStream());

				for (int sheet = 0; sheet < embeddedWorkbook.getNumberOfSheets(); sheet++) {
					final HSSFSheet activeSheet = embeddedWorkbook.getSheetAt(sheet);
					if (activeSheet.getSheetName().equalsIgnoreCase("Sheet1")) {
						for (int rowIndex = activeSheet.getFirstRowNum(); rowIndex <= activeSheet
								.getLastRowNum(); rowIndex++) {
							final HSSFRow row = activeSheet.getRow(rowIndex);
							for (int cellIndex = row.getFirstCellNum(); cellIndex <= row
									.getLastCellNum(); cellIndex++) {
								final HSSFCell cell = row.getCell(cellIndex);
								if (cell != null) {
									if (cell.getCellType() == Cell.CELL_TYPE_STRING)
										System.out.println("Row:" + rowIndex + " Cell:" + cellIndex + " = "
												+ cell.getStringCellValue());
									if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
										System.out.println("Row:" + rowIndex + " Cell:" + cellIndex + " = "
												+ cell.getNumericCellValue());

										cell.setCellValue(cell.getNumericCellValue() * 2); // update
																							// the
																							// value
									}
								}
							}
						}
					}
				}
			}
		}
	}
	
	
	@Test
	public void viewFile() {
		POIFSFileSystem fs = null;
		List strings = POIFSViewEngine.inspectViewable(fs, true, 0, "  ");
		Iterator iter = strings.iterator();

		while (iter.hasNext()) {
			//os.write( ((String)iter.next()).getBytes());
			System.out.println(iter.next());
		}
	}
}
