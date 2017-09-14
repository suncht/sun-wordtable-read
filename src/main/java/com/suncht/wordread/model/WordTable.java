package com.suncht.wordread.model;

import java.util.List;

import com.google.common.collect.Lists;
import com.suncht.wordread.format.IWordTableCellFormater;

/**
 * 一个表格对象
 * 每个表格有多个行
 * @author changtan.sun
 *
 */
public class WordTable {
	private List<WordTableRow> rows = Lists.newArrayList();

	public List<WordTableRow> getRows() {
		return rows;
	}

	@Override
	public String toString() {
		return rows.toString();
	}

	public String format(IWordTableCellFormater cellFormater) {
		StringBuilder builder = new StringBuilder();
		for (WordTableRow row : rows) {
			for (WordTableCell cell : row.getCells()) {
				builder.append(cellFormater.format(cell));
			}
			builder.append(this.newline());
		}

		return builder.toString();
	}

	private String newline() {
		return System.getProperty("line.separator");
	}

}
