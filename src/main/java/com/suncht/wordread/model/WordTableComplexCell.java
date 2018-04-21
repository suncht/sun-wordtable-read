package com.suncht.wordread.model;

/**
 * 复杂类的单元格
 * 单元格中嵌套子单元格，可以看做嵌入了表格
 * @author suncht
 *
 */
public class WordTableComplexCell extends WordTableCell {
	/**
	 * 单元格中嵌套子单元格，可以看做嵌入了表格
	 */
	private WordTable innerTable;

	public WordTable getInnerTable() {
		return innerTable;
	}

	public void setInnerTable(WordTable innerTable) {
		this.innerTable = innerTable;
	}

	@Override
	public String toString() {
		return innerTable.toString();
	}

}
