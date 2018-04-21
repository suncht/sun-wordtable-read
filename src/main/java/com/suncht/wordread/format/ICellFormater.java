package com.suncht.wordread.format;

import com.suncht.wordread.model.WordTableCellContent;

/**
 * 单元格数据格式化接口
 * @author changtan.sun
 *
 */
public interface ICellFormater {
	public Object format(WordTableCellContent cellContent);
	
}
