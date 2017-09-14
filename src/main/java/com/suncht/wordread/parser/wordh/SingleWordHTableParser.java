package com.suncht.wordread.parser.wordh;

import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableRow;

import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.ISingleWordTableParser;
import com.suncht.wordread.parser.WordTableMemoryMapping;
import com.suncht.wordread.parser.WordTableTransferContext;

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

	private void parseRow(TableRow row, int i) {
		// TODO Auto-generated method stub

	}

}
