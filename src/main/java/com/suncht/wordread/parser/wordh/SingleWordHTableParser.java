package com.suncht.wordread.parser.wordh;

import java.math.BigInteger;

import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableRow;

import com.google.common.base.Preconditions;
import com.suncht.wordread.model.TTCPr;
import com.suncht.wordread.model.TTCPr.TTCPrEnum;
import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.model.WordTableCellContents;
import com.suncht.wordread.parser.ISingleWordTableParser;
import com.suncht.wordread.parser.WordTableTransferContext;
import com.suncht.wordread.parser.mapping.WordTableMemoryMapping;

/**
 * Doc文档解析
 * 
 * <p>
 * 标题: SingleWordHTableParser
 * </p>
 * <p>
 * 描述: 对POI API进行调试发现，解析Doc单元格的方式与Docx方式不同：没有列合并，只有行合并，有列宽
 * </p>
 * 
 * @author changtan.sun
 * @date 2018年4月27日
 */
public class SingleWordHTableParser implements ISingleWordTableParser {
	private Table hwpfTable;

	private WordTableMemoryMapping _tableMemoryMapping;
	private WordTableTransferContext context;

	/**
	 * 最大列数
	 */
	private int realMaxColumnCount = 0;
	/**
	 * 最大列数所占的行Index
	 */
	private int rowIndexOfMaxColumnCount = 0;

	public SingleWordHTableParser(Table hwpfTable, WordTableTransferContext context) {
		this.hwpfTable = hwpfTable;
		this.context = context;
	}

	public WordTable parse() {
		int realMaxRowCount = this.hwpfTable.numRows();

		realMaxColumnCount = 0;
		for (int i = 0; i < realMaxRowCount; i++) {
			TableRow tr = this.hwpfTable.getRow(i);
			int numCell = tr.numCells();
			if (numCell > realMaxColumnCount) {
				realMaxColumnCount = numCell;
				rowIndexOfMaxColumnCount = i;
			}
		}

		_tableMemoryMapping = new WordTableMemoryMapping(realMaxRowCount, realMaxColumnCount);

		for (int i = 0; i < realMaxRowCount; i++) {
			TableRow preRow = i - 1 >= 0 ? this.hwpfTable.getRow(i - 1) : null; // 上一行
			parseRow(this.hwpfTable.getRow(i), i, preRow);
		}

		return context.transfer(_tableMemoryMapping);
	}

	private void parseRow(TableRow row, int realRowIndex, TableRow preRow) {
		int numCells = row.numCells();
		//boolean existColMergedCells = realMaxColumnCount > numCells; // 该行中是否存在被列合并，如果存在，做逻辑列合并处理
		int logicColumnIndex = 0;
		int logicRowIndex = realRowIndex; //逻辑行号和实际行号一样的
		for (int realColumnIndex = 0; realColumnIndex < numCells; realColumnIndex++) {
			TableCell cell = row.getCell(realColumnIndex);// 取得单元格
			int skipColumn = parseCell(row, cell, realRowIndex, realColumnIndex, logicRowIndex, logicColumnIndex);
			logicColumnIndex = logicColumnIndex + skipColumn + 1;
		}
	}

	/**
	 * 参考：https://blog.csdn.net/www1056481167/article/details/56835043
	 * 解析Doc单元格的方式与Docx方式不同：没有列合并概念，只有行合并
	 * 
	 * @param cell
	 * @param realRowIndex
	 * @param realColumnIndex
	 * @return
	 */
	private int parseCell(TableRow row, TableCell cell, int realRowIndex, int realColumnIndex, int logicRowIndex, int logicColumnIndex) {
		// -----列合并-----
		int numOfCellHMerged = computeNumOfCellHMerged(row, cell, realColumnIndex); //就是该单元格合并了多少列

		// -----行合并-----
		if (cell.isFirstVerticallyMerged() && cell.isVerticallyMerged()) { // 行合并开始
			TTCPr ttc = new TTCPr();
			if(numOfCellHMerged>0) {
				ttc.setType(TTCPrEnum.HVM_S); 
			} else {
				ttc.setType(TTCPrEnum.VM_S);
			}
			ttc.setRealRowIndex(realRowIndex);
			ttc.setRealColumnIndex(realColumnIndex);
			ttc.setLogicRowIndex(logicRowIndex);
			ttc.setLogicColumnIndex(logicColumnIndex);
			ttc.setWidth(BigInteger.valueOf(cell.getWidth()));
			ttc.setColSpan(numOfCellHMerged);
			ttc.setRoot(null);
			// ttc.setText(cell.getText());
			ttc.setContent(WordTableCellContents.getCellContent(cell));

			_tableMemoryMapping.setTTCPr(ttc, logicRowIndex, logicColumnIndex);
			
			//处理其他被合并的列
			if(numOfCellHMerged>0) {
				for (int i = 0; i < numOfCellHMerged; i++) {
					TTCPr ttc_merged = new TTCPr();
					ttc_merged.setType(TTCPrEnum.HM);
					ttc_merged.setRealRowIndex(realRowIndex);
					ttc_merged.setRealColumnIndex(realColumnIndex);
					ttc_merged.setLogicRowIndex(logicRowIndex);
					ttc_merged.setLogicColumnIndex(logicColumnIndex + i + 1);
					//ttc_merged.setWidth(BigInteger.valueOf(cell.getWidth()));
					//ttc_merged.setColSpan(numOfCellHMerged);
					ttc_merged.setRoot(ttc);
					
					_tableMemoryMapping.setTTCPr(ttc_merged, logicRowIndex, ttc_merged.getLogicColumnIndex());
				}
			}
		} else if (!cell.isFirstVerticallyMerged() && cell.isVerticallyMerged()) { // 行被合并
			int _start = logicRowIndex, _end = 0;
			TTCPr root = null;
			for (int i = logicRowIndex - 1; i >= 0; i--) {
				TTCPr ttcpr = _tableMemoryMapping.getTTCPr(i, logicColumnIndex);
				if (ttcpr != null && (ttcpr.getType() == TTCPrEnum.VM_S || ttcpr.getType() == TTCPrEnum.HVM_S)) {
					_end = i;
					root = ttcpr;
					break;
				} else if (ttcpr != null && ttcpr.getRoot() != null) {
					_end = i;
					root = ttcpr.getRoot();
					break;
				}
			}
			
			Preconditions.checkNotNull(root, "父单元格不能为空");

			TTCPr ttc = new TTCPr();
			ttc.setType(TTCPrEnum.VM);
			ttc.setRealRowIndex(realRowIndex);
			ttc.setRealColumnIndex(realColumnIndex);
			ttc.setLogicRowIndex(logicRowIndex);
			ttc.setLogicColumnIndex(logicColumnIndex);
			ttc.setWidth(BigInteger.valueOf(cell.getWidth()));
			ttc.setRoot(root);
			root.setRowSpan(_start - _end + 1);
			
			_tableMemoryMapping.setTTCPr(ttc, logicRowIndex, logicColumnIndex);
		} else { // 没有行合并
			TTCPr ttc = new TTCPr();
			if(numOfCellHMerged>0) {
				ttc.setType(TTCPrEnum.HM_S);
			} else {
				ttc.setType(TTCPrEnum.NONE);
			}
			ttc.setRealRowIndex(realRowIndex);
			ttc.setRealColumnIndex(realColumnIndex);
			ttc.setLogicRowIndex(logicRowIndex);
			ttc.setLogicColumnIndex(logicColumnIndex);
			ttc.setWidth(BigInteger.valueOf(cell.getWidth()));
			ttc.setColSpan(numOfCellHMerged);
			ttc.setRoot(null);
			// ttc.setText(cell.getText());
			ttc.setContent(WordTableCellContents.getCellContent(cell));

			_tableMemoryMapping.setTTCPr(ttc, logicRowIndex, logicColumnIndex);
			
			//处理其他被合并的列
			if(numOfCellHMerged>0) {
				for (int i = 0; i < numOfCellHMerged; i++) {
					TTCPr ttc_merged = new TTCPr();
					ttc_merged.setType(TTCPrEnum.HM);
					ttc_merged.setRealRowIndex(realRowIndex);
					ttc_merged.setRealColumnIndex(realColumnIndex);
					ttc_merged.setLogicRowIndex(logicRowIndex);
					ttc_merged.setLogicColumnIndex(logicColumnIndex + i + 1);
					//ttc_merged.setWidth(BigInteger.valueOf(cell.getWidth()));
					//ttc_merged.setColSpan(numOfCellHMerged);
					ttc_merged.setRoot(ttc);
					
					_tableMemoryMapping.setTTCPr(ttc_merged, logicRowIndex, ttc_merged.getLogicColumnIndex());
				}
			}
		}

		return numOfCellHMerged;
	}

	/**
	 * 计算合并了多少个单元格
	 * 表格中其他行根据标准行进行列合并，属于标准表格 标准表格，比如
	 * ——————————————— 
	 * |   |    |    |
	 * ——————————————— 
	 * | | |    |    | ---->该行为标准行 
	 * ——————————————— 
	 * | |           |
	 * ——————————————— 
	 * |        |    | 
	 * ———————————————
	 * 
	 * @param cell
	 * @param realRowIndex
	 * @param realColumnIndex
	 * @return
	 */
	private int computeNumOfCellHMerged(TableRow currentRow, TableCell currentCell, int realColumnIndex) {
		TableRow standardRow = this.hwpfTable.getRow(this.rowIndexOfMaxColumnCount);

		if (currentRow.numCells() >= standardRow.numCells()) {
			return 0;
		}

		long totalWidth = 0;
		for (int i = 0; i <= realColumnIndex; i++) {
			totalWidth += currentRow.getCell(i).getWidth();
		}

		int tempRowIndex = -1;
		long tempWidth = 0;
		for (int i = 0, size = standardRow.numCells(); i < size; i++) {
			tempWidth += standardRow.getCell(i).getWidth();
			if (this.widthEqual(tempWidth, totalWidth)) {
				tempRowIndex = i;
				break;
			}
		}
		
		int currentCellWidth = currentCell.getWidth();
		tempWidth = 0;
		int columnMerged = 0;
		for (int i = tempRowIndex; i >= 0; i--) {
			tempWidth += standardRow.getCell(i).getWidth();
			if(this.widthEqual(tempWidth, currentCellWidth)) {
				break;
			} else {
				columnMerged++;
			}
		}
		
		return columnMerged;
	}

	private boolean widthEqual(long tempWidth, long totalWidth) {
		return tempWidth <= (totalWidth + 10) && tempWidth >= (totalWidth - 10);
	}

}
