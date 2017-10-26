package com.suncht.wordread.parser.wordh;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.google.common.collect.Lists;
import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.ISingleWordTableParser;
import com.suncht.wordread.parser.IWordTableParser;
import com.suncht.wordread.parser.WordTableTransferContext;

public class WordHTableParser implements IWordTableParser {
	private WordTableTransferContext context;

	public WordHTableParser(WordTableTransferContext context) {
		this.context = context;
	}

	public List<WordTable> parse(InputStream inputStream) {

		List<WordTable> wordTables = Lists.newArrayList();

		try {
			POIFSFileSystem pfs = new POIFSFileSystem(inputStream); // 载入文档  
			HWPFDocument hwpf = new HWPFDocument(pfs);

			Range range = hwpf.getRange();//得到文档的读取范围  
			TableIterator it = new TableIterator(range);
			//迭代文档中的表格  
			while (it.hasNext()) {
				Table table = (Table) it.next();
				ISingleWordTableParser parser = new SingleWordHTableParser(table, context);
				WordTable wordTable = parser.parse();
				wordTables.add(wordTable);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

		return wordTables;
	}
}
