package com.suncht.wordread.model;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class WordTableCellContent {
	protected Object content;
	protected String oxml;

	public Object getContent() {
		return content;
	}

	public void setContent(Object content) {
		this.content = content;
	}

	public String getOxml() {
		return oxml;
	}

	protected void setOxml(String oxml) {
		this.oxml = oxml;
	}

	public WordTableCellContent copy() {
		WordTableCellContent newContent = new WordTableCellContent();
		newContent.setContent(content);
		newContent.setOxml(oxml);
		return newContent;
	}

	//	public static WordTableCellContent getCellContent(String oxml, String text) {
	//		WordTableCellContent content = null;
	//		if (oxml.contains("<m:oMathPara>") && oxml.contains("</m:oMathPara>")) { //是公式
	//			content = new WordTableCellContentOmml();
	//		} else { //一般文本
	//			content = new WordTableCellContentText();
	//		}
	//		content.setText(text);
	//		content.setOxml(oxml);
	//		return content;
	//	}

	public static WordTableCellContent getCellContent(XWPFTableCell cell) {
		WordTableCellContent content = null;
		if (isFormula(cell)) { //是公式
			content = new WordTableCellContentOmml(cell);
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
