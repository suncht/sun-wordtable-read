package com.test;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.junit.Test;

import com.suncht.wordread.format.DefaultWordTableCellFormater;
import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.WordTableParser;
import com.suncht.wordread.parser.WordTableParser.WordDocType;
import com.suncht.wordread.parser.strategy.LogicalTableStrategy;

/**
 * 嵌套图片单元格测试
 * @author suncht
 *
 */
public class NestedImageCellTest {
	@Test
	public void test01() throws IOException {
		InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/嵌套图片02.docx");
		List<WordTable> tables = WordTableParser.create().transferStrategy(new LogicalTableStrategy()).memoryMappingVisitor(new MemoryMappingVisitorTest()).parse(inputStream, WordDocType.DOCX);
		for (WordTable wordTable : tables) {
			System.out.println(wordTable.format(new DefaultWordTableCellFormater()));
		}

		inputStream.close();
	}
}
