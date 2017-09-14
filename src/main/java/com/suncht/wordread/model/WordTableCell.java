package com.suncht.wordread.model;

import com.suncht.wordread.format.IWordTableCellFormater;

/**
 * 表格的单元格
 * @author changtan.sun
 *
 */
public class WordTableCell {
	//	/**
	//	 * 行号
	//	 */
	//	private int rowIndex;
	//	/**
	//	 * 列号
	//	 */
	//	private int columnIndex;
	//	/**
	//	 * 在word中实际行号
	//	 */
	//	private int realRowIndex;
	//	/**
	//	 * 在word中实际列号
	//	 */
	//	private int realColumnIndex;
	/**
	 * 单元格的内容
	 */
	private String text;

	/**
	 * 在word中合并了多少行
	 */
	private int rowSpan = 1;
	/**
	 * 在word中合并了多少列
	 */
	private int columnSpan = 1;

	//	public String getCellPosition() {
	//		return rowIndex + "-" + columnIndex;
	//	}

	//	public int getRealRowIndex() {
	//		return realRowIndex;
	//	}
	//
	//	public void setRealRowIndex(int realRowIndex) {
	//		this.realRowIndex = realRowIndex;
	//	}
	//
	//	public int getRealColumnIndex() {
	//		return realColumnIndex;
	//	}
	//
	//	public void setRealColumnIndex(int realColumnIndex) {
	//		this.realColumnIndex = realColumnIndex;
	//	}

	public String getText() {
		return text;
	}

	public void setText(String text) {
		this.text = text;
	}

	public int getRowSpan() {
		return rowSpan;
	}

	public void setRowSpan(int rowSpan) {
		this.rowSpan = rowSpan;
	}

	public int getColumnSpan() {
		return columnSpan;
	}

	public void setColumnSpan(int columnSpan) {
		this.columnSpan = columnSpan;
	}

	//	public int getRowIndex() {
	//		return rowIndex;
	//	}
	//
	//	public void setRowIndex(int rowIndex) {
	//		this.rowIndex = rowIndex;
	//	}

	//	public int getColumnIndex() {
	//		return columnIndex;
	//	}

	//	public void setColumnIndex(int columnIndex) {
	//		this.columnIndex = columnIndex;
	//	}

	//	@Override
	//	public String toString() {
	//		return "PeraWordTableCell [rowIndex=" + rowIndex + ", columnIndex=" + columnIndex + ", text=" + text + "]";
	//	}

	/**
	 * 单元格数据格式化成字符串
	 * @param formater
	 * @return
	 */
	public String format(IWordTableCellFormater formater) {
		if (formater == null) {
			return this.toString();
		}

		return formater.format(this);
	}

}
