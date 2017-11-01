package com.suncht.wordread.parser.wordx;

import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import com.suncht.wordread.model.TTCPr;
import com.suncht.wordread.model.TTCPr.TTCPrEnum;
import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.model.WordTableCellContent;
import com.suncht.wordread.parser.ISingleWordTableParser;
import com.suncht.wordread.parser.WordTableTransferContext;
import com.suncht.wordread.parser.mapping.WordTableMemoryMapping;

/**
 * 解析Docx中一张复杂表格内容
 * @author changtan.sun
 *
 */
public class SingleWordXTableParser implements ISingleWordTableParser {
	private XWPFTable xwpfTable;
	//	private WordTable table;

	private WordTableMemoryMapping _tableMemoryMapping;
	private WordTableTransferContext context;

	public SingleWordXTableParser(XWPFTable xwpfTable, WordTableTransferContext context) {
		this.xwpfTable = xwpfTable;
		this.context = context;
	}

	//	public WordTable getTable() {
	//		return table;
	//	}

	/**
	 * 解析Docx的表格，将表格相关数据映射到表格映射对象中， 用于后面的操作
	 * @return
	 */
	public WordTable parse() {
		List<XWPFTableRow> rows;
		List<XWPFTableCell> cells;

		rows = xwpfTable.getRows();
		int realMaxRowCount = rows.size();
		//		table.setRealMaxRowCount(rows.size());

		int realMaxColumnCount = 0;
		for (XWPFTableRow row : rows) {
			//获取行对应的单元格  
			cells = row.getTableCells();
			if (cells.size() > realMaxColumnCount) {
				realMaxColumnCount = cells.size();
			}
		}

		//table.setRealMaxColumnCount(columnCount);

		_tableMemoryMapping = new WordTableMemoryMapping(realMaxRowCount, realMaxColumnCount);
		_tableMemoryMapping.setVisitor(context.getVisitor());
		for (int i = 0; i < realMaxRowCount; i++) {
			parseRow(rows.get(i), i);
		}

		//printTableMemoryMap();

		//		wordTableMap = new WordTableMap();
		//		wordTableMap.setTableMemoryMap(_tableMemoryMap);
		return context.transfer(_tableMemoryMapping);
	}

	public void dispose() {
		_tableMemoryMapping = null;
		xwpfTable = null;
	}

	//	/**
	//	 * 打印表格映射
	//	 */
	//	private void printTableMemoryMap() {
	//		int r = 1;
	//		for (TTCPr[] columns : _tableMemoryMapping) {
	//			int c = 1;
	//			for (TTCPr column : columns) {
	//				System.out.println(r + ":" + c + "===>" + column.getType() + " ==== " + column.getText());
	//				c++;
	//			}
	//
	//			r++;
	//		}
	//	}

	/**
	 * 解析word中表格行
	 * @param row
	 * @param realRowIndex
	 */
	private void parseRow(XWPFTableRow row, int realRowIndex) {
		List<XWPFTableCell> cells = row.getTableCells();
		//_tableMemoryMap[rowIndex] = new TTCPr[cells.size()];
		int columnIndex = 0;
		for (XWPFTableCell cell : cells) {
			//skipColumn是否跳过多个单元格, 当列合并时候
			int skipColumn = parseCell(cell, realRowIndex, columnIndex);
			columnIndex = columnIndex + skipColumn + 1;
		}
	}

	private int parseCell(XWPFTableCell cell, int realRowIndex, int realColumnIndex) {
		int skipColumn = 0;
		if (_tableMemoryMapping.getTTCPr(realRowIndex, realColumnIndex) != null) {
			return skipColumn;
		}

		CTTcPr tt = cell.getCTTc().getTcPr();
		String a = tt.xmlText();
		//-------行合并--------
		if (tt.getVMerge() != null) {
			if (tt.getVMerge().getVal() != null && "restart".equals(tt.getVMerge().getVal().toString())) { //行合并的第一行单元格(行合并的开始单元格)
				TTCPr ttc = new TTCPr();
				ttc.setType(TTCPrEnum.VM_S);
				ttc.setRealRowIndex(realRowIndex);
				ttc.setRealColumnIndex(realColumnIndex);
				ttc.setRoot(null);
				//ttc.setText(cell.getText());
				ttc.setContent(WordTableCellContent.getCellContent(cell));

				_tableMemoryMapping.setTTCPr(ttc, realRowIndex, realColumnIndex);
			} else { //行合并的其他行单元格（被合并的单元格）
				int _start = realRowIndex, _end = 0;
				TTCPr root = null;
				for (int i = realRowIndex - 1; i >= 0; i--) {
					TTCPr ttcpr = _tableMemoryMapping.getTTCPr(i, realColumnIndex);
					if (ttcpr != null && (ttcpr.getType() == TTCPrEnum.VM_S || ttcpr.getType() == TTCPrEnum.HVM_S)) {
						_end = i;
						root = _tableMemoryMapping.getTTCPr(_end, realColumnIndex);
						break;
					}
				}

				TTCPr ttc = new TTCPr();
				ttc.setType(TTCPrEnum.VM);
				ttc.setRealRowIndex(realRowIndex);
				ttc.setRealColumnIndex(realColumnIndex);
				ttc.setRoot(root);
				_tableMemoryMapping.setTTCPr(ttc, realRowIndex, realColumnIndex);
				root.setRowSpan(_start - _end + 1);
			}
		} else { //没有进行行合并的单元格
			TTCPr currentCell = _tableMemoryMapping.getTTCPr(realRowIndex, realColumnIndex);
			if (currentCell != null && currentCell.getType() == TTCPrEnum.HM) { //被列合并的单元格

			} else {
				currentCell = new TTCPr();
				currentCell.setType(TTCPrEnum.NONE);
				currentCell.setRealRowIndex(realRowIndex);
				currentCell.setRealColumnIndex(realColumnIndex);
				currentCell.setContent(WordTableCellContent.getCellContent(cell));
				currentCell.setRoot(null);

				//判断是否有父单元格
				if (realRowIndex > 0) {
					TTCPr parent = _tableMemoryMapping.getTTCPr(realRowIndex - 1, realColumnIndex);
					if (parent.isDoneColSpan()) {
						currentCell.setParent(parent);
					}
				}

				_tableMemoryMapping.setTTCPr(currentCell, realRowIndex, realColumnIndex);
			}
		}

		//-------列合并-------
		if (tt.getGridSpan() != null) {
			int colSpan = tt.getGridSpan().getVal().intValue();
			TTCPr root = _tableMemoryMapping.getTTCPr(realRowIndex, realColumnIndex);
			root.setColSpan(colSpan);
			if (root.getType() == TTCPrEnum.VM_S) {
				root.setType(TTCPrEnum.HVM_S);
			} else {
				root.setType(TTCPrEnum.HM_S);
			}

			//给其他被列合并的单元格进行初始化
			for (int i = 1; i < colSpan; i++) {
				TTCPr cell_other = _tableMemoryMapping.getTTCPr(realRowIndex, realColumnIndex + i);
				if (cell_other == null)
					cell_other = new TTCPr();
				cell_other.setRealRowIndex(realRowIndex);
				cell_other.setRealColumnIndex(realColumnIndex + i);
				cell_other.setType(TTCPrEnum.HM);
				cell_other.setRoot(root);

				_tableMemoryMapping.setTTCPr(cell_other, realRowIndex, realColumnIndex + i);
			}

			skipColumn = colSpan - 1;
		}

		return skipColumn;
	}
}
