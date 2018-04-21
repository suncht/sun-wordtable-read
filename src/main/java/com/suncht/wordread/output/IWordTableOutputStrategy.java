package com.suncht.wordread.output;

import com.suncht.wordread.model.WordTableCell;

/**
 * 表格单元格内容输出策略
 * @author suncht
 *
 */
public interface IWordTableOutputStrategy {
	public void output(WordTableCell tableCell);
}
