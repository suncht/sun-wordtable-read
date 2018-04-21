package com.suncht.wordread.format;

import com.suncht.wordread.model.ContentTypeEnum;
import com.suncht.wordread.model.WordTableCellContent;
import com.suncht.wordread.model.WordTableCellContentFormula;
import com.suncht.wordread.model.WordTableCellContentImage;
import com.suncht.wordread.model.WordTableCellContentImage.ImageContent;
import com.suncht.wordread.model.WordTableCellContentText;

/**
 * 单元格内容格式化的默认实现
 * @author suncht
 *
 */
public class DefaultCellFormater implements ICellFormater {
	@Override
	public Object format(WordTableCellContent cellContent) {
		if(cellContent.getContentType() == ContentTypeEnum.Text) {
			WordTableCellContentText _cellContent = (WordTableCellContentText)cellContent;
			return this.formatText(_cellContent);
		} else if(cellContent.getContentType() == ContentTypeEnum.Image) {
			WordTableCellContentImage _cellContent = (WordTableCellContentImage)cellContent;
			return this.formatImage(_cellContent);
		} else if(cellContent.getContentType() == ContentTypeEnum.Formula) {
			WordTableCellContentFormula _cellContent = (WordTableCellContentFormula)cellContent;
			return this.formatFormula(_cellContent);
		}
		return "";
	}

	public Object formatText(WordTableCellContentText cellContent) {
		String text = (String) cellContent.getData();
		return text;
	}
	
	public Object formatImage(WordTableCellContentImage cellContent) {
		ImageContent imageContent = (ImageContent)cellContent.getData();
		return imageContent!=null ? imageContent.getFileName(): "";
	}
	
	public Object formatFormula(WordTableCellContentFormula cellContent) {
		String formula = (String) cellContent.getData();
		return formula;
	}
}
