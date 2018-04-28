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

/**
 * Doc文档解析
 * 
* <p>标题: SingleWordHTableParser</p>  
* <p>描述: 对POI API进行调试发现，解析Doc单元格的方式与Docx方式不同：没有列合并，只有行合并，有列宽</p>  
* @author changtan.sun  
* @date 2018年4月27日
 */
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
			//System.out.println(cell.toString());
			int skipColumn = parseCell(cell, realRowIndex, i);
		}
	}

	/**
	 * 参考：https://blog.csdn.net/www1056481167/article/details/56835043
	 * 解析Doc单元格的方式与Docx方式不同：没有列合并概念，只有行合并
	 * @param cell
	 * @param realRowIndex
	 * @param realColumnIndex
	 * @return
	 */
	private int parseCell(TableCell cell, int realRowIndex, int realColumnIndex) {
		int a = cell.getStartOffset();
		int b = cell.getEndOffset();
		int c = cell.getLeftEdge();
		
		String text = cell.text().trim();
		System.out.println("文本:" +text + " === " + b + " ==== " + cell.isFirstVerticallyMerged() + " ==== " + cell.isVerticallyMerged() + " ==== " + cell.getWidth());
		System.out.println("row:" + realRowIndex + "   column:" + realColumnIndex);
		//System.out.println("文本:" +text + " ===> " + b + " ====> " + cell.getDescriptor().isFFirstMerged() + " ====> " + cell.getDescriptor().isFMerged());
		//System.out.println(cell.getDescriptor());
		
		int numParagraphs = cell.numParagraphs();
		for (int k = 0; k < numParagraphs; k++) {
			Paragraph para = cell.getParagraph(k);
			//System.out.println(para.getStartOffset() + " -- " + para.getEndOffset());
			String s = para.text();
			// 去除后面的特殊符号
			if (null != s && !"".equals(s)) {
				s = s.substring(0, s.length() - 1);
			}
			//System.out.println(s);
		}
		System.out.println("");
		return 0;
	}

}
