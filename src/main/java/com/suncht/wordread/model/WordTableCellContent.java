package com.suncht.wordread.model;

public abstract class WordTableCellContent {
	protected ContentTypeEnum contentType;
	protected Object data;
	

	public Object getData() {
		return data;
	}

	public void setData(Object data) {
		this.data = data;
	}
	
	public ContentTypeEnum getContentType() {
		return contentType;
	}

	public void setContentType(ContentTypeEnum contentType) {
		this.contentType = contentType;
	}

	public abstract WordTableCellContent copy();

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
}
