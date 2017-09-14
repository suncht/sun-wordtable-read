package com.suncht.wordread.parser;

import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.strategy.ITableTransferStrategy;

/**
 * Word表格转换上下文
 * @author changtan.sun
 *
 */
public class WordTableTransferContext {
	private ITableTransferStrategy strategy;

	public WordTableTransferContext(ITableTransferStrategy strategy) {
		this.strategy = strategy;
	}

	public WordTable transfer(final WordTableMemoryMapping tableMemoryMapping) {
		return strategy.transfer(tableMemoryMapping);
	}
}
