package com.test;

import com.suncht.wordread.model.TTCPr;
import com.suncht.wordread.parser.mapping.IWordTableMemoryMappingVisitor;

public class MemoryMappingVisitorTest implements IWordTableMemoryMappingVisitor {

	@Override
	public void visit(final TTCPr cell, int realRowIndex, int realColumnIndex) {
		if (realRowIndex == 0 && realColumnIndex == 0) {
			cell.setText("测试成功");
		}
	}

}
