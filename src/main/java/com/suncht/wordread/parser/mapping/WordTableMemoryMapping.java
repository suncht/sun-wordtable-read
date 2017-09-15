package com.suncht.wordread.parser.mapping;

import java.util.Arrays;

import com.google.common.base.Preconditions;
import com.suncht.wordread.model.TTCPr;

/**
 * Word表格内存映射
 * @author changtan.sun
 *
 */
public class WordTableMemoryMapping {
	private TTCPr[][] _tableMemoryMap;
	private int rowCount;
	private int columnCount;
	private IWordTableMemoryMappingVisitor visitor;

	public WordTableMemoryMapping(int row, int column) {
		_tableMemoryMap = new TTCPr[row][column];
		this.rowCount = row;
		this.columnCount = column;
	}

	public void setTTCPr(final TTCPr data, int rowIndex, int columnIndex) {
		Preconditions.checkArgument(rowIndex < rowCount);
		Preconditions.checkArgument(columnIndex < columnCount);

		_tableMemoryMap[rowIndex][columnIndex] = data;

		if (visitor != null) {
			data.accept(visitor, rowIndex, columnIndex);
		}
	}

	public final TTCPr getTTCPr(int rowIndex, int columnIndex) {
		Preconditions.checkArgument(rowIndex < rowCount);
		Preconditions.checkArgument(columnIndex < columnCount);

		return _tableMemoryMap[rowIndex][columnIndex];
	}

	public TTCPr[] getRow(int rowIndex) {
		Preconditions.checkArgument(rowIndex < rowCount);

		return Arrays.copyOf(_tableMemoryMap[rowIndex], columnCount);
	}

	public int getRowCount() {
		return rowCount;
	}

	public void setRowCount(int rowCount) {
		this.rowCount = rowCount;
	}

	public int getColumnCount() {
		return columnCount;
	}

	public void setColumnCount(int columnCount) {
		this.columnCount = columnCount;
	}

	public IWordTableMemoryMappingVisitor getVisitor() {
		return visitor;
	}

	public void setVisitor(IWordTableMemoryMappingVisitor visitor) {
		this.visitor = visitor;
	}

}
