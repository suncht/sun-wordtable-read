package com.suncht.wordread.format;

import com.suncht.wordread.model.WordTableCell;

/**
 * 表格数据格式化接口
 * @author changtan.sun
 *
 */
public interface IWordTableFormater {
	public void format(WordTableCell tableCell, StringBuilder builder);
}
