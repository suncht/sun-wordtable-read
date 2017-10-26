package com.suncht.wordread.model;

import java.util.List;

import com.google.common.collect.Lists;

/**
 * 表格行
 * 行中包含多个单元格
 * @author changtan.sun
 *
 */
public class WordTableRow {
	/**
	 * 行中单元格集合
	 */
	private List<WordTableCell> cells = Lists.newArrayList();

	public List<WordTableCell> getCells() {
		return cells;
	}

	@Override
	public String toString() {
		return cells.toString();
	}

	public void clear() {
		this.cells.clear();
		this.cells = null;
	}
}
