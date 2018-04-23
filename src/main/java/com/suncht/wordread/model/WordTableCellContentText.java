package com.suncht.wordread.model;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class WordTableCellContentText extends WordTableCellContent {
	private final static Logger logger = LoggerFactory.getLogger(WordTableCellContentText.class);
	
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
