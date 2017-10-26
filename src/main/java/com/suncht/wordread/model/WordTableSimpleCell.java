package com.suncht.wordread.model;

public class WordTableSimpleCell extends WordTableCell {

	@Override
	public String toString() {
		return getContent().getText();
	}

}
