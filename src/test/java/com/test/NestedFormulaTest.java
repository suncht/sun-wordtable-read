package com.test;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.junit.Test;

import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.WordTableParser;
import com.suncht.wordread.parser.WordTableParser.WordDocType;
import com.suncht.wordread.parser.strategy.LogicalTableStrategy;

public class NestedFormulaTest {
	@Test
	public void testFormulaInCell() throws IOException {
		try(InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/嵌套公式.docx");) {
			List<WordTable> tables = WordTableParser.create().transferStrategy(new LogicalTableStrategy())
					.memoryMappingVisitor(new MemoryMappingVisitorTest()).parse(inputStream, WordDocType.DOCX);
			for (WordTable wordTable : tables) {
				System.out.println(wordTable.format());
			}
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
}
