package com.suncht.wordread.model;

import java.io.StringReader;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.InputSource;

import com.suncht.wordread.utils.MathmlUtils;


public class WordTableCellContentFormula extends WordTableCellContent {
	private final static Logger logger = LoggerFactory.getLogger(WordTableCellContentFormula.class);
	
	private String oxml;
	
	public WordTableCellContentFormula() {

	}

	public String getOxml() {
		return oxml;
	}

	protected void setOxml(String oxml) {
		this.oxml = oxml;
	}

	
	public WordTableCellContentFormula(XWPFTableCell cell) {
		String xml = cell.getCTTc().xmlText();
		String omml = this.extractOml(xml);

		String mml = MathmlUtils.convertOMML2MML(omml);
		String latex = MathmlUtils.convertMML2Latex(mml);
		this.setData(latex);
		this.setOxml(xml);
		this.setContentType(ContentTypeEnum.Formula);
	}

	@Override
	public WordTableCellContent copy() {
		WordTableCellContentFormula newContent = new WordTableCellContentFormula();
		newContent.setData(this.data);
		newContent.setOxml(this.oxml);
		newContent.setContentType(ContentTypeEnum.Formula);
		return newContent;
	}

	private String extractOml(String xml) {
		//dom4j解析器的初始化
		SAXReader reader = new SAXReader(new DocumentFactory());
		Map<String, String> map = new HashMap<String, String>();
		map.put("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
		map.put("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
		reader.getDocumentFactory().setXPathNamespaceURIs(map); //xml文档的namespace设置

		InputSource source = new InputSource(new StringReader(xml));
		source.setEncoding("utf-8");
		try {
			Document doc = reader.read(source);
			Element root = doc.getRootElement();
			Element e = (Element) root.selectSingleNode("//m:oMathPara"); //用xpath得到OMML节点
			String omml = e.asXML(); //转为xml
			return omml;
		} catch (DocumentException e) {
			e.printStackTrace();
		}
		return null;
	}

}
