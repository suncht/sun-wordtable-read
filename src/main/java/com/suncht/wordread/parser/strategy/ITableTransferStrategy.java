package com.suncht.wordread.parser.strategy;

import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.mapping.WordTableMemoryMapping;

/**
 * 表格转换策略
 * 将表格内存映射转换成实际的表格模式
 * @author changtan.sun
 *
 */
public interface ITableTransferStrategy {
	public WordTable transfer(WordTableMemoryMapping tableMemoryMapping);
}
