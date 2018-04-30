package com.suncht.wordread.parser.strategy;

import com.suncht.wordread.model.TTCPr;
import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.model.WordTableCell;
import com.suncht.wordread.model.WordTableComplexCell;
import com.suncht.wordread.model.WordTableRow;
import com.suncht.wordread.model.WordTableSimpleCell;
import com.suncht.wordread.parser.mapping.WordTableMemoryMapping;

public class LogicalTableStrategy implements ITableTransferStrategy {

	private WordTableMemoryMapping tableMemoryMapping;

	//	/**
	//	 * 获取在docx中实际行数（word中表格都处理成二维表格，忽略合并）
	//	 * @return
	//	 */
	//	private int getRealMaxRowCount() {
	//		return tableMemoryMap.length;
	//	}

	/**
	 * 获取行数（在表格映射对象中的行数）
	 * @return
	 */
	private int getRowCount() {
		int rowCount = 0;
		for (int i = 0; i < tableMemoryMapping.getRowCount(); i++) {
			if (tableMemoryMapping.getTTCPr(i, 0).isValid()) {
				rowCount++;
			}
		}
		return rowCount;
	}

	public WordTable transfer(WordTableMemoryMapping tableMemoryMapping) {
		this.tableMemoryMapping = tableMemoryMapping;

		WordTable wordTable = new WordTable();
		int rowCount = getRowCount();
		WordTableRow tableRow = null;
		for (int i = 0; i < rowCount; i++) {
			tableRow = this.getTableRow(i);
			wordTable.getRows().add(tableRow);
		}
		return wordTable;
	}

	/**
	 * 获取行对象
	 * @param currentRowIndex
	 * @return
	 */
	private WordTableRow getTableRow(int currentRowIndex) {
		TTCPr[] _rows = null;
		TTCPr _first_column_in_row = null;
		int rowCount = 0;
		for (int i = 0; i < tableMemoryMapping.getRowCount(); i++) {
			if (tableMemoryMapping.getTTCPr(i, 0).isValid()) {
				if (currentRowIndex == rowCount++) {
					_rows = tableMemoryMapping.getRow(i);
					_first_column_in_row = tableMemoryMapping.getTTCPr(i, 0);
					break;
				}
			}
		}

		if (_rows == null) {
			return null;
		}

		int _logic_row_index = _first_column_in_row.getLogicRowIndex();
		//int _end_row_index = _first_column_in_row.getRowSpan() + _first_column_in_row.getRealRowIndex() - 1;
		int _row_span = _first_column_in_row.getRowSpan();
		int _logic_column_count = _rows.length;

		WordTableRow pwtr = new WordTableRow();

		WordTableCell cell = null;
		for (int i = 0; i < _logic_column_count; i++) {
			cell = getCellInRow(_logic_row_index, _row_span, i, currentRowIndex);
			if (cell == null) {
				continue;
			}
			pwtr.getCells().add(cell);
		}

		return pwtr;
	}

	/**
	 * 获取一行中的单元格集合,将实际单元格转换成逻辑单元格
	 * @param logicRowIndex 逻辑行号
	 * @param endRealRowIndex 逻辑行号
	 * @param logicColumnIndex  word中的实际列
	 * @param currentRowIndex  在表格映射对象中的行号
	 * @return
	 */
	private WordTableCell getCellInRow(int logicRowIndex, int logicRowSpan, int logicColumnIndex, int currentRowIndex) {
		WordTableCell cell = null;
		TTCPr currentRealCell = tableMemoryMapping.getTTCPr(logicRowIndex, logicColumnIndex);

		boolean needHandleRowSpan = logicRowSpan > 0 || currentRealCell.isDoneRowSpan(); //是否需要处理跨行的情况
		boolean needHandleColSpan = currentRealCell.isDoneColSpan();//是否需要处理跨列的情况

		boolean satisfyConditionOfComplexCell = false; //是否满足复杂单元格的条件

		satisfyConditionOfComplexCell = needHandleRowSpan && needHandleColSpan;
		if (!satisfyConditionOfComplexCell) {
			satisfyConditionOfComplexCell = currentRealCell.getRowSpan() < logicRowSpan;
		}

		if (currentRealCell.isValid()) { //有效单元格
			if (satisfyConditionOfComplexCell) {//跨行又跨列
				WordTableComplexCell pwtc = new WordTableComplexCell(); //属于复杂单元格

				WordTable innerTable = new WordTable();
				int _realColSpan = currentRealCell.getColSpan();
				for (int i = 0; i < logicRowSpan;) {
					WordTableRow innerRow = new WordTableRow();
					int _rowSpan = 1;
					for (int j = 0; j < _realColSpan; j++) {
						TTCPr _ttcpr = tableMemoryMapping.getTTCPr(logicRowIndex + i, logicColumnIndex + j);
						if (_ttcpr.isValid()) {
							WordTableCell _cell = new WordTableSimpleCell();
							_cell.setRowSpan(_ttcpr.getRowSpan());
							_cell.setColumnSpan(_ttcpr.getColSpan());
							_cell.setContent(_ttcpr.getContent().copy());
							innerRow.getCells().add(_cell);

							if (_ttcpr.getRowSpan() > _rowSpan) {
								_rowSpan = _ttcpr.getRowSpan();
							}
						}
					}
					innerTable.getRows().add(innerRow);

					i = i + _rowSpan;
				}
				pwtc.setInnerTable(innerTable);
				cell = pwtc;
			} else {
				//跨列不跨行，不需要处理
				//跨行不跨列，不需要处理
				WordTableSimpleCell pwtc = new WordTableSimpleCell(); //属于简单单元格
				pwtc.setRowSpan(currentRealCell.getRowSpan());
				pwtc.setColumnSpan(currentRealCell.getColSpan());
				pwtc.setContent(currentRealCell.getContent().copy());

				cell = pwtc;
			}
		}

		return cell;

	}
}
