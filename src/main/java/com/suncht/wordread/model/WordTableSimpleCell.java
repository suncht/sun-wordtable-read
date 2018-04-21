package com.suncht.wordread.model;

/**
 * 简单单元格
 * 比如：文字、公式、附件等
 * @author suncht
 *
 */
public class WordTableSimpleCell extends WordTableCell {

	@Override
	public String toString() {
		return getContent().getData().toString();
	}

}
