package com.suncht.wordread.format;

import java.util.List;

import com.suncht.wordread.model.WordTableCell;
import com.suncht.wordread.model.WordTableCellContent;
import com.suncht.wordread.model.WordTableComplexCell;
import com.suncht.wordread.model.WordTableRow;
import com.suncht.wordread.model.WordTableSimpleCell;

/**
 * 默认的单元格内容Formatter
 * 
 * @author suncht
 *
 */
public class DefaultWordTableFormater implements IWordTableFormater {
	private ICellFormater cellFormater;

	private StringBuilder builder;

	public DefaultWordTableFormater() {
		this.cellFormater = new DefaultCellFormater();
	}

	public DefaultWordTableFormater(ICellFormater cellFormater) {
		this.cellFormater = cellFormater;
	}

	public void format(WordTableCell tableCell, StringBuilder builder) {
		this.builder = builder!=null ? builder : new StringBuilder();
		
		if (tableCell instanceof WordTableSimpleCell) {
			printCell(tableCell.getContent());
		} else if (tableCell instanceof WordTableComplexCell) {
			WordTableComplexCell cell = (WordTableComplexCell) tableCell;

			List<WordTableRow> rows = cell.getInnerTable().getRows();
			for (WordTableRow row : rows) {
				for (WordTableCell wtcell : row.getCells()) {
					printCell(wtcell.getContent());
				}
			}
		}
	}

	private void printCell(WordTableCellContent cellContent) {
		Object data = cellFormater.format(cellContent);
		builder.append(data.toString());
	}
}
