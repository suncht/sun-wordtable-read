package com.suncht.wordread.model;

import java.io.StringReader;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.DocumentFactory;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.dom4j.tree.DefaultElement;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.InputSource;

import com.suncht.wordread.model.WordTableCellContentImage.WcImage;
import com.suncht.wordread.parser.WordTableParser.WordDocType;

public class WordTableCellContentImage extends WordTableCellContent<WcImage> {
	private final static Logger logger = LoggerFactory.getLogger(WordTableCellContentImage.class);
	
	public WordTableCellContentImage(WordDocType docType) {
		this.docType = docType;
	}
	
	@Override
	public void load(Object cellObj) {
		this.setContentType(ContentTypeEnum.Image);
		
		if(docType == WordDocType.DOCX) {
			XWPFTableCell cell = (XWPFTableCell) cellObj;
			String xml = cell.getCTTc().xmlText();
			String embedId = extractEmbedId(xml);
			this.setData(this.readImage(embedId, cell.getXWPFDocument()));
		} else if(docType == WordDocType.DOC) {
			
		}
	}
	
	
	@Override
	public WordTableCellContent<WcImage> copy() {
		WordTableCellContentImage newContent = new WordTableCellContentImage(this.docType);
		newContent.setData(data);
		newContent.setContentType(contentType);
		return newContent;
	}

	private String extractEmbedId(String xml) {
		// dom4j解析器的初始化
		SAXReader reader = new SAXReader(new DocumentFactory());
		Map<String, String> map = new HashMap<String, String>();
		map.put("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
		map.put("a", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
		map.put("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
		map.put("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
		map.put("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
		map.put("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
		reader.getDocumentFactory().setXPathNamespaceURIs(map); // xml文档的namespace设置

		InputSource source = new InputSource(new StringReader(xml));
		source.setEncoding("utf-8");
		try {
			Document doc = reader.read(source);
			Element root = doc.getRootElement();
			Element e = (Element) root.selectSingleNode("//pic:blipFill");
			Element blip = (DefaultElement) e.content().get(0);
			String embedId = ((Attribute) (blip.attributes().get(0))).getValue();
			return embedId;
		} catch (DocumentException e) {
			e.printStackTrace();
		}
		return null;
	}

	private WcImage readImage(String embedId, final XWPFDocument xdoc) {
		if (StringUtils.isBlank(embedId)) {
			return null;
		}
		WcImage imageContent = null;
		for (XWPFPictureData pictureData : xdoc.getAllPictures()) {
			PackageRelationship relationship = pictureData.getPackageRelationship();
			if (embedId.equals(relationship.getId())) {
				imageContent = new WcImage();
				imageContent.setData(pictureData.getData());
				imageContent.setFileName(pictureData.getFileName());
				imageContent.setImageType(pictureData.getPictureType());
				break;
			}
		}

		return imageContent;
	}

	/**
	 * 图片内容
	* <p>标题: ImageContent</p>  
	* <p>描述: </p>  
	* @author changtan.sun  
	* @date 2018年4月22日
	 */
	public static class WcImage {
		private String fileName;
		private byte[] data;
		
		/**
		 * 图片类型，参考org.apache.poi.xwpf.usermodel
		 */
		private int imageType;

		public String getFileName() {
			return fileName;
		}

		public void setFileName(String fileName) {
			this.fileName = fileName;
		}

		public byte[] getData() {
			if (data == null) {
				return new byte[0];
			}
			return Arrays.copyOf(data, data.length);
		}

		public void setData(byte[] data) {
			if (data == null) {
				return;
			}
			this.data = Arrays.copyOf(data, data.length);
		}

		public int getImageType() {
			return imageType;
		}

		public void setImageType(int imageType) {
			this.imageType = imageType;
		}

	}


}
