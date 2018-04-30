package com.suncht.wordread.model;

import com.suncht.wordread.parser.WordTableParser.WordDocType;

/**
* <p>标题: 单元格内容对象</p>  
* <p>描述: </p>  
* @author changtan.sun  
* @date 2018年4月22日
 */
public abstract class WordTableCellContent<T> {
	protected WordDocType docType;
	
	protected ContentTypeEnum contentType;
	protected T data;

	public T getData() {
		return data;
	}

	public void setData(T data) {
		this.data = data;
	}
	
	public ContentTypeEnum getContentType() {
		return contentType;
	}

	public void setContentType(ContentTypeEnum contentType) {
		this.contentType = contentType;
	}

	/**
	 * 拷贝对象，具体实现由子类实现
	 * @return
	 */
	public abstract WordTableCellContent<T> copy();
	
	public abstract void load(Object cellObj);
}
