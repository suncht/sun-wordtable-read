package com.suncht.wordread.model;

import java.io.IOException;
import java.io.StringReader;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.dom4j.tree.DefaultElement;
import org.xml.sax.InputSource;

import com.sun.xml.internal.txw2.output.StreamSerializer;

/**
* <p>标题: 单元格中嵌套OLE对象</p>  
* <p>描述: OLE对象，比如：附件</p>  
* @author changtan.sun  
* @date 2018年4月22日
 */
public class WordTableCellContentOleObject extends WordTableCellContent {

	public WordTableCellContentOleObject() {
		
	}
	
	public WordTableCellContentOleObject(XWPFTableCell cell) {
		String xml = cell.getCTTc().xmlText();
		Document doc = this.buildDocument(xml);
		String embedId = extractOleObjectEmbedId(doc);
		
		OleObjectContent oleObjectContent = this.readOleObject(embedId, cell.getXWPFDocument());
		this.setData(oleObjectContent);
		this.setContentType(ContentTypeEnum.OleObject);
	}
	
	@Override
	public WordTableCellContent copy() {
		WordTableCellContentOleObject newContent = new WordTableCellContentOleObject();
		newContent.setData(data);
		newContent.setContentType(contentType);
		return newContent;
	}
	
	/**
	 * 由单元格内容xml构建Document
	 * @param xml
	 * @return
	 */
	private Document buildDocument(String xml) {
		// dom4j解析器的初始化
		SAXReader reader = new SAXReader(new DocumentFactory());
		Map<String, String> map = new HashMap<String, String>();
		map.put("o", "urn:schemas-microsoft-com:office:office");
		map.put("v", "urn:schemas-microsoft-com:vml");
		map.put("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
		map.put("a", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
		map.put("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
		map.put("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
		map.put("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		reader.getDocumentFactory().setXPathNamespaceURIs(map); // xml文档的namespace设置

		InputSource source = new InputSource(new StringReader(xml));
		source.setEncoding("utf-8");
		
		try {
			Document doc = reader.read(source);
			return doc;
		} catch (DocumentException e) {
			e.printStackTrace();
		}
		return null;
	}
	
	/**
	 * 从单元格Document中获取OLE对象的embedId
	 * @param doc
	 * @return
	 */
	private String extractOleObjectEmbedId(Document doc) {
		Element root = doc.getRootElement();
		Element e = (Element) root.selectSingleNode("//w:object");
		Element oOLEObject = (DefaultElement) e.content().get(1);
		String embedId = ((Attribute) (oOLEObject.attribute("id"))).getValue();
		return embedId;
	}
	
	/**
	 * 从单元格Document中获取附件的显示图片的embedId
	 * @param doc
	 * @return
	 */
	private String extractImageEmbedId(Document doc) {
		Element root = doc.getRootElement();
		Element e = (Element) root.selectSingleNode("//w:object");
		Element vShape = (DefaultElement) e.content().get(0);
		Element vImagedata = (Element) vShape.selectSingleNode("//v:imagedata");
		String embedId = ((Attribute) (vImagedata.attribute("id"))).getValue();
		return embedId;
	}
	
	/**
	 * 读取Ole对象
	 * @param embedId
	 * @param xdoc
	 * @return
	 */
	private OleObjectContent readOleObject(String embedId, final XWPFDocument xdoc) {
		if (StringUtils.isBlank(embedId)) {
			return null;
		}
		OleObjectContent oleObjectContent = null;
		List<POIXMLDocumentPart> parts = xdoc.getRelations();
		for (POIXMLDocumentPart poixmlDocumentPart : parts) {
			String id = poixmlDocumentPart.getPackageRelationship().getId();
			if(embedId.equals(id)) {
				PackagePart packagePart = poixmlDocumentPart.getPackagePart();
				
				oleObjectContent = new OleObjectContent();
				oleObjectContent.setFileName(packagePart.getPartName().getName());
				
				try {
					byte[] data = IOUtils.toByteArray(packagePart.getInputStream());
					oleObjectContent.setData(data);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

		return oleObjectContent;
	}
	
	public static class OleObjectContent {
		private String fileName;
		private byte[] data;

		
		public String getFileName() {
			return fileName;
		}

		public void setFileName(String fileName) {
			this.fileName = fileName;
		}

		public byte[] getData() {
			return data;
		}

		public void setData(byte[] data) {
			this.data = data;
		}
	}

}
