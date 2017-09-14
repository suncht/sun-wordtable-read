package com.test;

import java.io.InputStream;
import java.util.List;

import org.junit.Test;

import com.suncht.wordread.format.DefaultWordTableCellFormater;
import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.WordTableParser;
import com.suncht.wordread.parser.WordTableParser.WordDocType;
import com.suncht.wordread.parser.strategy.LogicalTableStrategy;

public class WordHTableParserTest {
	@Test
	public void test01() {
		InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/1.doc");
		//InputStream inputStream = new FileInputStream(new File(doc2));
		List<WordTable> tables = WordTableParser.create().transferStrategy(new LogicalTableStrategy()).parse(inputStream, WordDocType.DOC);
		for (WordTable wordTable : tables) {
			System.out.println(wordTable.format(new DefaultWordTableCellFormater()));
		}
	}
}
