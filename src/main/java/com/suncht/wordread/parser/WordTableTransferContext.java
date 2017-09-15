package com.suncht.wordread.parser;

import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.mapping.IWordTableMemoryMappingVisitor;
import com.suncht.wordread.parser.mapping.WordTableMemoryMapping;
import com.suncht.wordread.parser.strategy.ITableTransferStrategy;
import com.suncht.wordread.parser.strategy.LogicalTableStrategy;

/**
 * Word表格转换上下文
 * @author changtan.sun
 *
 */
public class WordTableTransferContext {
	private ITableTransferStrategy strategy;
	private IWordTableMemoryMappingVisitor visitor;

	public static WordTableTransferContext create() {
		return new WordTableTransferContext();
	}

	public WordTableTransferContext transferStrategy(ITableTransferStrategy strategy) {
		this.strategy = strategy;
		return this;
	}

	public WordTableTransferContext visitor(IWordTableMemoryMappingVisitor visitor) {
		this.visitor = visitor;
		return this;
	}

	public WordTable transfer(final WordTableMemoryMapping tableMemoryMapping) {
		if (strategy == null) {
			strategy = new LogicalTableStrategy();
		}
		return strategy.transfer(tableMemoryMapping);
	}

	public ITableTransferStrategy getStrategy() {
		return strategy;
	}

	public IWordTableMemoryMappingVisitor getVisitor() {
		return visitor;
	}

}
