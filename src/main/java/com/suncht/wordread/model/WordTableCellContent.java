package com.suncht.wordread.model;

/**
* <p>标题: 单元格内容对象</p>  
* <p>描述: </p>  
* @author changtan.sun  
* @date 2018年4月22日
 */
public abstract class WordTableCellContent {
	protected ContentTypeEnum contentType;
	protected Object data;

	public Object getData() {
		return data;
	}

	public void setData(Object data) {
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
	public abstract WordTableCellContent copy();
}
