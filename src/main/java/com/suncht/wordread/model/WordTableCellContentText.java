package com.suncht.wordread.model;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.suncht.wordread.model.WordTableCellContentText.WcText;
import com.suncht.wordread.parser.WordTableParser.WordDocType;

public class WordTableCellContentText extends WordTableCellContent<WcText> {
	private final static Logger logger = LoggerFactory.getLogger(WordTableCellContentText.class);
	
	public WordTableCellContentText(WordDocType docType) {
		this.docType = docType;
	}
	
	@Override
	public void load(Object cellObj) {
		this.setContentType(ContentTypeEnum.Text);
		
		if(docType == WordDocType.DOCX) {
			XWPFTableCell cell = (XWPFTableCell) cellObj;
			List<String> texts = new ArrayList<String>();
			
			List<XWPFParagraph> paragraphs = cell.getParagraphs();
			if(paragraphs!=null && paragraphs.size()>0) {
				for (XWPFParagraph paragraph : paragraphs) {
					texts.add(this.runsToLine(paragraph.getRuns()));
				}
			} 
			
			WcText text = new WcText();
			text.setParagraphs(texts);
			
			this.setData(text);
		} else if(docType == WordDocType.DOC) {
			TableCell cell = (TableCell) cellObj;
			
			List<String> texts = new ArrayList<String>();
			
			for (int i = 0, num = cell.numParagraphs(); i < num; i++) {
				Paragraph paragraph = cell.getParagraph(i);
				texts.add(paragraph.text().trim());
			}
			
			WcText text = new WcText();
			text.setParagraphs(texts);
			
			this.setData(text);
		}
	}
	
	private String runsToLine(List<XWPFRun> runs) {
		StringBuilder builder = new StringBuilder();
		for (XWPFRun run : runs) {
			builder.append(run.toString());
		}
		
		return builder.toString();
	}
	
	public WordTableCellContent<WcText> copy() {
		WordTableCellContentText newContent = new WordTableCellContentText(this.docType);
		newContent.setData(data);
		newContent.setContentType(contentType);
		return newContent;
	}
	
	/**
	 * 文本
	* <p>标题: WcText</p>  
	* <p>描述: </p>  
	* @author changtan.sun  
	* @date 2018年4月23日
	 */
	public static class WcText {
		private List<String> paragraphs;

		public List<String> getParagraphs() {
			return paragraphs;
		}

		public void setParagraphs(List<String> paragraphs) {
			this.paragraphs = Collections.unmodifiableList(paragraphs);
		}

		@Override
		public String toString() {
			return StringUtils.join(paragraphs, '\n');
		}
		
		
	}
}
