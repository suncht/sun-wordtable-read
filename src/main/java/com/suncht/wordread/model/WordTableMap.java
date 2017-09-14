package com.suncht.wordread.model;

/**
 * 表格数据映射
 * @author changtan.sun
 *
 */
public class WordTableMap {
	private TTCPr[][] tableMemoryMap;

	public TTCPr[][] getTableMemoryMap() {
		return tableMemoryMap;
	}

	public void setTableMemoryMap(TTCPr[][] tableMemoryMap) {
		this.tableMemoryMap = tableMemoryMap;
	}

	public void clear() {
		tableMemoryMap = null;
	}

	/**
	 * 获取在docx中实际行数（word中表格都处理成二维表格，忽略合并）
	 * @return
	 */
	public int getRealMaxRowCount() {
		return tableMemoryMap.length;
	}

	/**
	 * 获取行数（在表格映射对象中的行数）
	 * @return
	 */
	public int getRowCount() {
		int rowCount = 0;
		for (int i = 0; i < tableMemoryMap.length; i++) {
			if (tableMemoryMap[i][0].isValid()) {
				rowCount++;
			}
		}
		return rowCount;
	}

	/**
	 * 获取行对象
	 * @param currentRowIndex
	 * @return
	 */
	public WordTableRow getRow(int currentRowIndex) {
		TTCPr[] _rows = null;
		TTCPr _first_column_in_row = null;
		int rowCount = 0;
		for (int i = 0; i < tableMemoryMap.length; i++) {
			if (tableMemoryMap[i][0].isValid()) {
				if (currentRowIndex == rowCount++) {
					_rows = tableMemoryMap[i];
					_first_column_in_row = tableMemoryMap[i][0];
					break;
				}
			}
		}

		if (_rows == null) {
			return null;
		}

		int real_row_index = _first_column_in_row.getRealRowIndex();
		//int _end_row_index = _first_column_in_row.getRowSpan() + _first_column_in_row.getRealRowIndex() - 1;
		int _row_span = _first_column_in_row.getRowSpan();
		int _real_column_count = _rows.length;

		WordTableRow pwtr = new WordTableRow();

		WordTableCell cell = null;
		//		WordTableCell pwtc = null;
		for (int i = 0; i < _real_column_count; i++) {
			cell = getCellInRow(real_row_index, _row_span, i, currentRowIndex);
			if (cell == null) {
				continue;
			}
			pwtr.getCells().add(cell);
			//			if (cells.size() == 1) {
			//				pwtr.getCells().add(cells.get(0));
			//			} else {
			//				pwtc = new WordTableCell();
			//				pwtc.getSubCells().addAll(cells);
			//				pwtr.getCells().add(pwtc);
			//			}
		}

		return pwtr;
	}

	/**
	 * 获取一行中的单元格集合,将实际单元格转换成逻辑单元格
	 * @param realRowIndex word中的实际开始行号
	 * @param endRealRowIndex word中的实际结束行号
	 * @param realColumnIndex  word中的实际列
	 * @param currentRowIndex  在表格映射对象中的行号
	 * @return
	 */
	private WordTableCell getCellInRow(int realRowIndex, int realRowSpan, int realColumnIndex, int currentRowIndex) {
		WordTableCell cell = null;
		TTCPr currentRealCell = tableMemoryMap[realRowIndex][realColumnIndex];

		boolean needHandleRowSpan = realRowSpan > 1 || currentRealCell.isDoneRowSpan(); //是否需要处理跨行的情况
		boolean needHandleColSpan = currentRealCell.isDoneColSpan();//是否需要处理跨列的情况

		boolean satisfyConditionOfComplexCell = false; //是否满足复杂单元格的条件

		satisfyConditionOfComplexCell = needHandleRowSpan && needHandleColSpan;
		if (!satisfyConditionOfComplexCell) {
			satisfyConditionOfComplexCell = currentRealCell.getRowSpan() < realRowSpan;
		}

		if (currentRealCell.isValid()) { //有效单元格
			if (satisfyConditionOfComplexCell) {//跨行又跨列
				WordTableComplexCell pwtc = new WordTableComplexCell(); //属于复杂单元格

				WordTable innerTable = new WordTable();
				int _realColSpan = currentRealCell.getColSpan();
				for (int i = 0; i < realRowSpan;) {
					WordTableRow innerRow = new WordTableRow();
					int _rowSpan = 1;
					for (int j = 0; j < _realColSpan; j++) {
						TTCPr _ttcpr = tableMemoryMap[realRowIndex + i][realColumnIndex + j];
						if (_ttcpr.isValid()) {
							WordTableCell _cell = new WordTableSimpleCell();
							_cell.setRowSpan(_ttcpr.getRowSpan());
							_cell.setColumnSpan(_ttcpr.getColSpan());
							_cell.setText(_ttcpr.getText());
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
				pwtc.setText(currentRealCell.getText());

				cell = pwtc;
			}
		}

		return cell;

		//		if (currentCell.isValid()) { //有效单元格
		//			pwtc = new WordTableCell();
		//			//				pwtc.setRealColumnIndex(realColumnIndex);
		//			//				pwtc.setRealRowIndex(i);
		//			//				pwtc.setColumnSpan(pttcpr.getColSpan());
		//			//				pwtc.setRowSpan(pttcpr.getRowSpan());
		//			//			pwtc.setRowIndex(currentRowIndex);
		//			//			pwtc.setColumnIndex(realColumnIndex);
		//			pwtc.setText(currentCell.getText());
		//
		//			if (currentCell.getType() == TTCPrEnum.VM_S) {
		//
		//			} else if (currentCell.getType() == TTCPrEnum.HM_S) {
		//
		//			} else if (currentCell.getType() == TTCPrEnum.HVM_S) {
		//
		//			}
		//
		//			cells.put(currentCell.getCellPosition(), pwtc);
		//		} else { //无效单元格
		//			if (i == realRowIndex) { //如果第一个单元格就是无效单元格， 当行合并时
		//				if (currentCell.getType() == TTCPrEnum.VM && currentCell.getRoot() != null) {
		//					TTCPr root = currentCell.getRoot();
		//					pwtc = new WordTableCell();
		//					//						pwtc.setRealColumnIndex(root.getRealColumnIndex());
		//					//						pwtc.setColumnSpan(root.getColSpan());
		//					//						pwtc.setRealRowIndex(i);
		//					//						pwtc.setRowSpan(root.getRowSpan());
		//					//					pwtc.setRowIndex(currentRowIndex);
		//					//					pwtc.setColumnIndex(realColumnIndex);
		//					pwtc.setText(currentCell.getText());
		//
		//					cells.put(pwtc.getCellPosition(), pwtc);
		//				}
		//			}
		//
		//			if (currentCell.getType() == TTCPrEnum.HM && currentCell.getRoot() != null) { //被行合并
		//				pwtc = new WordTableCell();
		//				//					pwtc.setRealColumnIndex(realColumnIndex);
		//				//					pwtc.setColumnSpan(pttcpr.getColSpan());
		//				//					pwtc.setRealRowIndex(i);
		//				//					pwtc.setRowSpan(pttcpr.getRowSpan());
		//				//				pwtc.setRowIndex(currentRowIndex);
		//				//				pwtc.setColumnIndex(realColumnIndex);
		//				pwtc.setText(currentCell.getText());
		//
		//				cells.add(pwtc);
		//			}
		//		}
		//
		//		for (int i = realRowIndex; i <= endRealRowIndex; i++) {
		//			currentCell = tableMemoryMap[i][realColumnIndex];
		//
		//		}
		//
		//		return cells;
	}
}
