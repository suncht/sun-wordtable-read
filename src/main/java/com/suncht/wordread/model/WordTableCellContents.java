package com.suncht.wordread.model;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public final class WordTableCellContents {
	public static WordTableCellContent getCellContent(XWPFTableCell cell) {
		WordTableCellContent content = null;
		if (isFormula(cell)) { //是公式
			content = new WordTableCellContentFormula(cell);
		} else if (isImage(cell)) { //图片
			content = new WordTableCellContentImage(cell);
		} else { //一般文本
			content = new WordTableCellContentText(cell);
		}
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

}
