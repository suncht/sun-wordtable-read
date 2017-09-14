package com.suncht.wordread.parser;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

import org.springframework.util.StringUtils;

import com.suncht.wordread.model.WordTable;
import com.suncht.wordread.parser.strategy.ITableTransferStrategy;
import com.suncht.wordread.parser.strategy.LogicalTableStrategy;
import com.suncht.wordread.parser.wordh.WordHTableParser;
import com.suncht.wordread.parser.wordx.WordXTableParser;

/**
 * Word文档解析器
 * 支持2007以上的docx、2007以下的doc文档
 * @author changtan.sun
 *
 */
public class WordTableParser {
	private static final String DOCX_WORD_DOCUMENT = ".docx";
	private static final String DOC_WORD_DOCUMENT = ".doc";

	private WordTableTransferContext context;
	private IWordTableParser wordTableParser;

	private WordTableParser() {

	}

	public static WordTableParser create() {
		return new WordTableParser();
	}

	public WordTableParser transferStrategy(ITableTransferStrategy tableTransferStrategy) {
		this.context = new WordTableTransferContext(tableTransferStrategy);
		return this;
	}

	public List<WordTable> parse(String wordFilePath) {
		if (this.context == null) {
			this.context = new WordTableTransferContext(new LogicalTableStrategy());
		}
		WordDocType docType = WordDocType.DOCX;
		if (StringUtils.endsWithIgnoreCase(wordFilePath, DOCX_WORD_DOCUMENT)) {
			docType = WordDocType.DOCX;
		} else if (StringUtils.endsWithIgnoreCase(wordFilePath, DOC_WORD_DOCUMENT)) {
			docType = WordDocType.DOC;
		} else {
			throw new IllegalArgumentException("不支持该文件类型");
		}

		try {
			FileInputStream inputStream = new FileInputStream(wordFilePath);
			return this.parse(inputStream, docType);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		return null;
	}

	public List<WordTable> parse(InputStream inputStream, WordDocType docType) {
		if (this.context == null) {
			this.context = new WordTableTransferContext(new LogicalTableStrategy());
		}

		if (docType == WordDocType.DOCX) {
			wordTableParser = new WordXTableParser(this.context);
		} else if (docType == WordDocType.DOC) {
			wordTableParser = new WordHTableParser(this.context);
		} else {
			throw new IllegalArgumentException("不支持该文件类型");
		}
		return wordTableParser.parse(inputStream);
	}

	/**
	 * Word文档类型
	 * @author changtan.sun
	 *
	 */
	public static enum WordDocType {
		DOCX, DOC, UNKOWN
	}
}
