package com.suncht.wordread.output;

import com.suncht.wordread.format.DefaultCellFormater;
import com.suncht.wordread.format.ICellFormater;
import com.suncht.wordread.model.WordTableCell;
import com.suncht.wordread.model.WordTableCellContentImage.ImageContent;
import com.suncht.wordread.model.WordTableComplexCell;
import com.suncht.wordread.model.WordTableSimpleCell;

public class DefaultWordTableOutputStrategy implements IWordTableOutputStrategy {
	private ICellFormater cellFormater;
	
	public DefaultWordTableOutputStrategy() {
		cellFormater = new DefaultCellFormater();
	}
	
	public DefaultWordTableOutputStrategy(ICellFormater cellFormater) {
		this.cellFormater = cellFormater;
	}

	@Override
	public void output(WordTableCell tableCell) {
		if (tableCell instanceof WordTableSimpleCell) {
			outputCell(tableCell.getContent().getData());
		}  else if (tableCell instanceof WordTableComplexCell) {
//			WordTableComplexCell cell = (WordTableComplexCell) tableCell;
//
//			StringBuilder builder = new StringBuilder();
//
//			List<WordTableRow> rows = cell.getInnerTable().getRows();
//			for (WordTableRow row : rows) {
//				for (WordTableCell wtcell : row.getCells()) {
//					builder.append(printCell(wtcell.getContent().getData()) + '\t');
//				}
//			}
//			return builder.toString() + "" + '\t';
		}
	}
	

	private void outputCell(Object cellContent) {
//		if (cellContent instanceof ImageContent) {
//			this.cellFormater.formatImage((ImageContent)cellContent);
//		} else {
//			this.cellFormater.formatText(cellContent);
//		}
	}

}
