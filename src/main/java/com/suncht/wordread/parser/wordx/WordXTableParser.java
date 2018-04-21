package com.suncht.wordread.parser.wordx;

import java.io.InputStream;
import java.util.List;

import org.apache.commons.io.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.collect.Lists;
import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.ISingleWordTableParser;
import com.suncht.wordread.parser.IWordTableParser;
import com.suncht.wordread.parser.WordTableTransferContext;

/**
 * Docx文档的复杂表格解析器
 * @author changtan.sun
 *
 */
public class WordXTableParser implements IWordTableParser {
	private final static Logger logger = LoggerFactory.getLogger(WordXTableParser.class);
	
	private WordTableTransferContext context;

	public WordXTableParser(WordTableTransferContext context) {
		this.context = context;
	}

	public List<WordTable> parse(InputStream inputStream) {
		List<WordTable> wordTables = Lists.newArrayList();

		try {
			XWPFDocument doc = new XWPFDocument(inputStream); // 载入文档  

			//获取文档中所有的表格  
			List<XWPFTable> tables = doc.getTables();
			for (XWPFTable table : tables) {
				ISingleWordTableParser parser = new SingleWordXTableParser(table, this.context);
				WordTable wordTable = parser.parse();
				wordTables.add(wordTable);
			}
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
		} finally {
			IOUtils.closeQuietly(inputStream);
		}

		return wordTables;
	}
}
