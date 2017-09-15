package com.suncht.wordread.parser.mapping;

import com.suncht.wordread.model.TTCPr;

/**
 * Word表格内存映射表的单元格访问者接口
 * 用于修改内存映射表的单元格的数据
 * @author changtan.sun
 *
 */
public interface IWordTableMemoryMappingVisitor {
	public void visit(TTCPr cell, int realRowIndex, int realColumnIndex);
}
