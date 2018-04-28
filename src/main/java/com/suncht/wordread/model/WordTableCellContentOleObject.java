package com.suncht.wordread.model;

import java.io.InputStream;
import java.io.StringReader;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.dev.POIFSViewEngine;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentNode;
import org.apache.poi.poifs.filesystem.Entry;
import org.apache.poi.poifs.filesystem.Ole10Native;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
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

import com.suncht.wordread.model.WordTableCellContentOleObject.WcOleObject;

/**
 * <p>
 * 标题: 单元格中嵌套OLE对象
 * </p>
 * <p>
 * 描述: OLE对象，比如：附件
 * </p>
 * 
 * @author changtan.sun
 * @date 2018年4月22日
 */
public class WordTableCellContentOleObject extends WordTableCellContent<WcOleObject> {
	private final static Logger logger = LoggerFactory.getLogger(WordTableCellContentOleObject.class);

	public WordTableCellContentOleObject() {

	}

	public WordTableCellContentOleObject(XWPFTableCell cell) {
		String xml = cell.getCTTc().xmlText();
		Document doc = this.buildDocument(xml);
		String embedId = extractOleObjectEmbedId(doc);

		WcOleObject oleObject = this.readOleObject(embedId, cell.getXWPFDocument());
		this.setData(oleObject);
		this.setContentType(ContentTypeEnum.OleObject);
	}

	@Override
	public WordTableCellContent<WcOleObject> copy() {
		WordTableCellContentOleObject newContent = new WordTableCellContentOleObject();
		newContent.setData(data);
		newContent.setContentType(contentType);
		return newContent;
	}

	/**
	 * 由单元格内容xml构建Document
	 * 
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
			logger.error(e.getMessage(), e);
		}
		return null;
	}

	/**
	 * 从单元格Document中获取OLE对象的embedId
	 * 
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
	 * 
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
	 * 
	 * @param embedId
	 * @param xdoc
	 * @return
	 */
	private WcOleObject readOleObject(String embedId, final XWPFDocument xdoc) {
		if (StringUtils.isBlank(embedId)) {
			return null;
		}
		WcOleObject oleObject = null;
		List<POIXMLDocumentPart> parts = xdoc.getRelations();
		for (POIXMLDocumentPart poixmlDocumentPart : parts) {
			String id = poixmlDocumentPart.getPackageRelationship().getId();
			if (embedId.equals(id)) {
				PackagePart packagePart = poixmlDocumentPart.getPackagePart();

				oleObject = new WcOleObject();
				// oleObjectContent.setFileName(packagePart.getPartName().getName());

				// 解析Ole对象中的文件，参考：http://poi.apache.org/poifs/how-to.html
				try (InputStream is = packagePart.getInputStream();) {
					POIFSFileSystem poifs = new POIFSFileSystem(is);

					if (isOle10NativeObject(poifs.getRoot())) {
						oleObject = readOle10Native(poifs);
					} else {
						oleObject = readDocumentOle(poifs, is);
					}
				} catch (Exception e) {
					logger.error(e.getMessage(), e);
				}
			}
		}

		return oleObject;
	}

	private boolean isOle10NativeObject(DirectoryNode directory) {
		return directory.hasEntry(Ole10Native.OLE10_NATIVE);
	}

	/**
	 * 读取非文档类的Ole对象
	 * 
	 * @param poifs
	 * @return
	 */
	private WcOleObject readOle10Native(POIFSFileSystem poifs) {
		WcOleObject oleObject = new WcOleObject();
		try {
			Ole10Native ole10 = Ole10Native.createFromEmbeddedOleObject(poifs);
			oleObject.setFileName(FilenameUtils.getName(ole10.getFileName()));

			// byte[] data = IOUtils.toByteArray(packagePart.getInputStream());
			oleObject.setDataSize(ole10.getDataSize());
			oleObject.setData(ole10.getDataBuffer());
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
		}

		return oleObject;
	}

	/**
	 * 读取文档类的OLE对象，包括Docx、Doc、xlsx、xls、ppt、pptx等
	 * 暂未实现，需要进行大量资料查阅、寻找解决方案
	 * @param poifs
	 * @return
	 */
	private WcOleObject readDocumentOle(POIFSFileSystem poifs, InputStream is) {
		DirectoryNode directory = poifs.getRoot();
		if (!directory.hasEntry("WordDocument")) {
			return null;
		}
		
		List strings = POIFSViewEngine.inspectViewable(poifs, true, 0, "  ");
		Iterator iter = strings.iterator();

		while (iter.hasNext()) {
			//os.write( ((String)iter.next()).getBytes());
			System.out.println(iter.next());
		}
		throw new NotImplementedException("暂未实现");
		
		
//		WcOleObject oleObject = new WcOleObject();
//		try {
//			DocumentNode entry = (DocumentNode) directory.getEntry("WpsCustomData");
//			byte[] data = new byte[entry.getSize()];
//			directory.createDocumentInputStream(entry).read(data);
//			
//			XWPFDocument doc = new XWPFDocument(directory.createDocumentInputStream(entry)); // 载入文档  
//			doc.toString();
//			oleObject.setFileName(FilenameUtils.getName(entry.getName()));
//			//
//			// //byte[] data =
//			// IOUtils.toByteArray(packagePart.getInputStream());
//			oleObject.setDataSize(data.length);
//			oleObject.setData(data);
//		} catch (Exception e) {
//			logger.error(e.getMessage(), e);
//		}

		//return oleObject;
	}

	public static class WcOleObject {
		private String fileName;
		private byte[] data;
		private int dataSize;

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

		public int getDataSize() {
			return dataSize;
		}

		public void setDataSize(int dataSize) {
			this.dataSize = dataSize;
		}
	}

}
