package com.suncht.wordread.parser.wordh;

import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableCellDescriptor;
import org.apache.poi.hwpf.usermodel.TableRow;

import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.ISingleWordTableParser;
import com.suncht.wordread.parser.WordTableTransferContext;
import com.suncht.wordread.parser.mapping.WordTableMemoryMapping;

public class SingleWordHTableParser implements ISingleWordTableParser {
	private Table hwpfTable;

	private WordTableMemoryMapping _tableMemoryMapping;
	private WordTableTransferContext context;

	public SingleWordHTableParser(Table hwpfTable, WordTableTransferContext context) {
		this.hwpfTable = hwpfTable;
		this.context = context;
	}

	public WordTable parse() {
		int realMaxRowCount = this.hwpfTable.numRows();

		int realMaxColumnCount = 0;
		for (int i = 0; i < realMaxRowCount; i++) {
			TableRow tr = this.hwpfTable.getRow(i);
			int numCell = tr.numCells();
			if (numCell > realMaxColumnCount) {
				realMaxColumnCount = numCell;
			}
		}

		_tableMemoryMapping = new WordTableMemoryMapping(realMaxRowCount, realMaxColumnCount);

		for (int i = 0; i < realMaxRowCount; i++) {
			parseRow(this.hwpfTable.getRow(i), i);
		}

		return context.transfer(_tableMemoryMapping);
	}

	private void parseRow(TableRow row, int realRowIndex) {
		int numCells = row.numCells();
		for (int i = 0; i < numCells; i++) {
			TableCell cell = row.getCell(i);// 取得单元格

			int skipColumn = parseCell(cell, realRowIndex, i);
		}
	}

	private int parseCell(TableCell cell, int realRowIndex, int realColumnIndex) {
		int a = cell.getStartOffset();
		int b = cell.getEndOffset();
		int c = cell.getLeftEdge();
		System.out.println(a + " === " + b + " ==== " + cell.isVerticallyMerged() + " ==== " + cell.isFirstVerticallyMerged());
		System.out.println(a + " ===> " + b + " ====> " + cell.getDescriptor().isFFirstMerged() + " ====> " + cell.getDescriptor().isFMerged());
		System.out.println(cell.getDescriptor());
		
		int numParagraphs = cell.numParagraphs();
		for (int k = 0; k < numParagraphs; k++) {
			Paragraph para = cell.getParagraph(k);
			System.out.println(para.getStartOffset() + " -- " + para.getEndOffset());
			String s = para.text();
			// 去除后面的特殊符号
			if (null != s && !"".equals(s)) {
				s = s.substring(0, s.length() - 1);
			}
			System.out.println(s);
		}
		System.out.println("");
		return 0;
	}

}
