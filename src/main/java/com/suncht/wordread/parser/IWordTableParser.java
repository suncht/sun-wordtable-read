package com.suncht.wordread.parser;

import java.io.InputStream;
import java.util.List;

import com.suncht.wordread.model.WordTable;

public interface IWordTableParser {
	public List<WordTable> parse(InputStream inputStream);
}
