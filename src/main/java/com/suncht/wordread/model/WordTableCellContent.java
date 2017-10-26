package com.suncht.wordread.model;

public class WordTableCellContent {
	protected String text;
	protected String oxml;

	public String getText() {
		return text;
	}

	protected void setText(String text) {
		this.text = text;
	}

	public String getOxml() {
		return oxml;
	}

	protected void setOxml(String oxml) {
		this.oxml = oxml;
	}

	public WordTableCellContent copy() {
		WordTableCellContent newContent = new WordTableCellContent();
		newContent.setText(text);
		newContent.setOxml(oxml);
		return newContent;
	}

	public static WordTableCellContent getCellContent(String oxml, String text) {
		WordTableCellContent content = null;
		if (oxml.contains("<m:oMathPara>") && oxml.contains("</m:oMathPara>")) { //是公式
			content = new WordTableCellContentOmml();
		} else { //一般文本
			content = new WordTableCellContentText();
		}
		content.setText(text);
		content.setOxml(oxml);
		return content;
	}

}
