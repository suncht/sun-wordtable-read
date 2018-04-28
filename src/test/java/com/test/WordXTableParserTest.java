package com.test;

import java.io.InputStream;
import java.util.List;

import org.junit.Test;

import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.WordTableParser;
import com.suncht.wordread.parser.WordTableParser.WordDocType;
import com.suncht.wordread.parser.strategy.LogicalTableStrategy;

public class WordXTableParserTest {
	String doc1 = "D:\\故障模式分析表格样例01.docx";
	String doc2 = "D:\\故障模式分析表格样例.docx";

	@Test
	public void test01() {
		try (InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/1.docx");) {
			List<WordTable> tables = WordTableParser.create().transferStrategy(new LogicalTableStrategy())
					.memoryMappingVisitor(new MemoryMappingVisitorTest()).parse(inputStream, WordDocType.DOCX);
			for (WordTable wordTable : tables) {
				System.out.println(wordTable.format());
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test
	public void test02() {
		InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/故障模式分析表格样例.docx");
		// InputStream inputStream = new FileInputStream(new File(doc2));
		List<WordTable> tables = WordTableParser.create().transferStrategy(new LogicalTableStrategy())
				.parse(inputStream, WordDocType.DOCX);
		for (WordTable wordTable : tables) {
			System.out.println(wordTable.format());
		}
	}

	@Test
	public void test03() {
		InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/故障模式分析表格样例01.docx");
		// InputStream inputStream = new FileInputStream(new File(doc2));
		List<WordTable> tables = WordTableParser.create().transferStrategy(new LogicalTableStrategy())
				.parse(inputStream, WordDocType.DOCX);
		for (WordTable wordTable : tables) {
			System.out.println(wordTable.format());
		}
	}
	
	@Test
	public void test04() {
		InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/复杂表格.docx");
		// InputStream inputStream = new FileInputStream(new File(doc2));
		List<WordTable> tables = WordTableParser.create().transferStrategy(new LogicalTableStrategy())
				.parse(inputStream, WordDocType.DOCX);
		for (WordTable wordTable : tables) {
			System.out.println(wordTable.format());
		}
	}
}
