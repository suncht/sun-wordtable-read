package com.suncht.wordread.model;

import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import com.suncht.wordread.parser.WordTableParser.WordDocType;

public final class WordTableCellContents {
	public static WordTableCellContent<?> getCellContent(XWPFTableCell cell) {
		WordTableCellContent<?> content = null;
		if (isFormula(cell)) { //是公式
			content = new WordTableCellContentFormula(WordDocType.DOCX);
		} else if (isImage(cell)) { //图片
			content = new WordTableCellContentImage(WordDocType.DOCX);
		} else if (isOleObject(cell)) { //OLE对象
			content = new WordTableCellContentOleObject(WordDocType.DOCX);
		} else { //一般文本
			content = new WordTableCellContentText(WordDocType.DOCX);
		}
		
		content.load(cell);
		return content;
	}

	public static boolean isFormula(XWPFTableCell cell) {
		String xmlText = cell.getCTTc().xmlText();
		return xmlText.contains("<m:oMathPara>") && xmlText.contains("</m:oMathPara>");
	}

	public static boolean isImage(XWPFTableCell cell) {
		String xmlText = cell.getCTTc().xmlText();
		return xmlText.contains("<w:drawing>") && xmlText.contains("</w:drawing>");
	}
	
	public static boolean isOleObject(XWPFTableCell cell) {
		String xmlText = cell.getCTTc().xmlText();
		return xmlText.contains("<w:object>") && xmlText.contains("</w:object>");
	}
	
	public static WordTableCellContent<?> getCellContent(TableCell cell) {
		WordTableCellContent<?> content = null;
		if (isFormula(cell)) { //是公式
			content = new WordTableCellContentFormula(WordDocType.DOC);
		} else if (isImage(cell)) { //图片
			content = new WordTableCellContentImage(WordDocType.DOC);
		} else if (isOleObject(cell)) { //OLE对象
			content = new WordTableCellContentOleObject(WordDocType.DOC);
		} else { //一般文本
			content = new WordTableCellContentText(WordDocType.DOC);
		}
		
		content.load(cell);
		return content;
	}
	
	public static boolean isFormula(TableCell cell) {
		String xmlText = cell.text();
		return xmlText.contains("<m:oMathPara>") && xmlText.contains("</m:oMathPara>");
	}
	
	public static boolean isImage(TableCell cell) {
		String xmlText = cell.text();
		return xmlText.contains("<w:drawing>") && xmlText.contains("</w:drawing>");
	}
	
	public static boolean isOleObject(TableCell cell) {
		String xmlText = cell.text();
		return xmlText.contains("<w:object>") && xmlText.contains("</w:object>");
	}

}
