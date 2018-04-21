package com.suncht.wordread.model;

import java.util.List;

import com.google.common.collect.Lists;
import com.suncht.wordread.format.DefaultWordTableFormater;
import com.suncht.wordread.format.IWordTableFormater;
import com.suncht.wordread.output.IWordTableOutputStrategy;

/**
 * 一个表格对象 每个表格有多个行
 * 
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

	public void output(IWordTableOutputStrategy outputStrategy) {
		for (WordTableRow row : rows) {
			for (WordTableCell cell : row.getCells()) {
				outputStrategy.output(cell);
			}
		}
	}

	public String format(IWordTableFormater tableFormater) {
		if (tableFormater == null) {
			tableFormater = new DefaultWordTableFormater();
		}

		StringBuilder builder = new StringBuilder();
		for (WordTableRow row : rows) {
			for (WordTableCell cell : row.getCells()) {
				tableFormater.format(cell, builder);
			}
			builder.append(this.newline());
		}

		return builder.toString();
	}

	public String format() {
		return this.format(null);
	}

	private String newline() {
		return System.getProperty("line.separator");
	}

}
