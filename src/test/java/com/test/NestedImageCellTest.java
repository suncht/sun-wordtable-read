package com.test;

import java.io.InputStream;
import java.util.List;

import org.junit.Test;

import com.suncht.wordread.format.DefaultCellFormater;
import com.suncht.wordread.format.DefaultWordTableFormater;
import com.suncht.wordread.format.IWordTableFormater;
import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.output.DefaultWordTableOutputStrategy;
import com.suncht.wordread.output.IWordTableOutputStrategy;
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
	public void test01() {
		IWordTableFormater tableFormater = new DefaultWordTableFormater(new DefaultCellFormater());
		IWordTableOutputStrategy outputStrategy = new DefaultWordTableOutputStrategy();
		
		try(InputStream inputStream = WordXTableParserTest.class.getResourceAsStream("/嵌套图片02.docx");) {
			List<WordTable> tables = WordTableParser.create().transferStrategy(new LogicalTableStrategy()).parse(inputStream, WordDocType.DOCX);
			
			for (WordTable wordTable : tables) {
				System.out.println(wordTable.format(tableFormater));
				wordTable.output(outputStrategy);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
}
