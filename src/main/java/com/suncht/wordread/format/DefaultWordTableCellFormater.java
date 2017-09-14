package com.suncht.wordread.format;

import java.util.List;

import com.suncht.wordread.model.WordTableCell;
import com.suncht.wordread.model.WordTableComplexCell;
import com.suncht.wordread.model.WordTableRow;
import com.suncht.wordread.model.WordTableSimpleCell;

public class DefaultWordTableCellFormater implements IWordTableCellFormater {
	public String format(WordTableCell tableCell) {
		if (tableCell instanceof WordTableSimpleCell) {
			return tableCell.getText() + '\t';
		} else if (tableCell instanceof WordTableComplexCell) {
			WordTableComplexCell cell = (WordTableComplexCell) tableCell;

			StringBuilder builder = new StringBuilder();

			List<WordTableRow> rows = cell.getInnerTable().getRows();
			for (WordTableRow row : rows) {
				builder.append(row.getCells().toString());
			}
			return builder.toString() + "" + '\t';
		}

		return "";
	}

}
