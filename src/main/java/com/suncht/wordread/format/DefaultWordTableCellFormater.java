package com.suncht.wordread.format;

import java.util.List;

import com.suncht.wordread.model.WordTableCell;
import com.suncht.wordread.model.WordTableCellContentImage.ImageContent;
import com.suncht.wordread.model.WordTableComplexCell;
import com.suncht.wordread.model.WordTableRow;
import com.suncht.wordread.model.WordTableSimpleCell;

public class DefaultWordTableCellFormater implements IWordTableCellFormater {
	public String format(WordTableCell tableCell) {
		if (tableCell instanceof WordTableSimpleCell) {
			return printCell(tableCell.getContent().getData()) + '\t';
		} else if (tableCell instanceof WordTableComplexCell) {
			WordTableComplexCell cell = (WordTableComplexCell) tableCell;

			StringBuilder builder = new StringBuilder();

			List<WordTableRow> rows = cell.getInnerTable().getRows();
			for (WordTableRow row : rows) {
				for (WordTableCell wtcell : row.getCells()) {
					builder.append(printCell(wtcell.getContent().getData()) + '\t');
				}
			}
			return builder.toString() + "" + '\t';
		}

		return "";
	}

	private String printCell(Object cellContent) {
		if (cellContent instanceof ImageContent) {
			return ((ImageContent) cellContent).getFileName();
		} else {
			return cellContent.toString();
		}

	}
}
