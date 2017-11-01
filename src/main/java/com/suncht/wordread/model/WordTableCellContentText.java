package com.suncht.wordread.model;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;

public class WordTableCellContentText extends WordTableCellContent {
	public WordTableCellContentText(XWPFTableCell cell) {
		this.setContent(cell.getText());
		this.setOxml(cell.getCTTc().xmlText());
	}
}
