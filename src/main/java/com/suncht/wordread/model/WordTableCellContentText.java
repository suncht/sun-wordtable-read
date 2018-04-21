package com.suncht.wordread.model;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class WordTableCellContentText extends WordTableCellContent {
	public WordTableCellContentText() {
		
	}
	
	public WordTableCellContentText(XWPFTableCell cell) {
		this.setData(cell.getText());
		this.setContentType(ContentTypeEnum.Text);
	}
	
	public WordTableCellContent copy() {
		WordTableCellContentText newContent = new WordTableCellContentText();
		newContent.setData(data);
		newContent.setContentType(contentType);
		return newContent;
	}
}
