package com.suncht.wordread.format;

import com.suncht.wordread.model.WordTableCell;

/**
 * 单元格数据格式化接口
 * @author changtan.sun
 *
 */
public interface IWordTableCellFormater {
	public String format(WordTableCell tableCell);
}
