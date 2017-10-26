package com.test;

import java.io.InputStream;
import java.util.List;

import org.junit.Test;

import com.suncht.wordread.format.DefaultWordTableCellFormater;
import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.WordTableParser;
import com.suncht.wordread.parser.WordTableParser.WordDocType;
import com.suncht.wordread.parser.strategy.LogicalTableStrategy;

public class WordCellDataTest {
	@Test
	public void testFormulaInCell() {
		InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/3.docx");
		List<WordTable> tables = WordTableParser.create().transferStrategy(new LogicalTableStrategy()).memoryMappingVisitor(new MemoryMappingVisitorTest()).parse(inputStream, WordDocType.DOCX);
		for (WordTable wordTable : tables) {
			System.out.println(wordTable.format(new DefaultWordTableCellFormater()));
		}
	}
}
