package com.suncht.wordread.format;

import com.suncht.wordread.model.ContentTypeEnum;
import com.suncht.wordread.model.WordTableCellContent;
import com.suncht.wordread.model.WordTableCellContentFormula;
import com.suncht.wordread.model.WordTableCellContentImage;
import com.suncht.wordread.model.WordTableCellContentImage.WcImage;
import com.suncht.wordread.model.WordTableCellContentOleObject;
import com.suncht.wordread.model.WordTableCellContentOleObject.WcOleObject;
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
		} else if(cellContent.getContentType() == ContentTypeEnum.OleObject) {
			WordTableCellContentOleObject _cellContent = (WordTableCellContentOleObject)cellContent;
			return this.formatOleObject(_cellContent);
		}
		return "";
	}


	public Object formatText(WordTableCellContentText cellContent) {
		String text = cellContent.getData().toString();
		return text;
	}
	
	public Object formatImage(WordTableCellContentImage cellContent) {
		WcImage imageContent = (WcImage)cellContent.getData();
		return imageContent!=null ? imageContent.getFileName(): "";
	}
	
	public Object formatFormula(WordTableCellContentFormula cellContent) {
		String formula = cellContent.getData().getLatex();
		return formula;
	}
	
	private Object formatOleObject(WordTableCellContentOleObject cellContent) {
		WcOleObject oleObject = cellContent.getData();
		return oleObject!=null ? oleObject.getFileName(): "";
	}
}
