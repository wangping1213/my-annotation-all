package com.wp.excel;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 抽象Excel2007读取器，excel2007的底层数据结构是xml文件，采用SAX的事件驱动的方法解析
 * xml，需要继承DefaultHandler，在遇到文件内容时，事件会触发，这种做法可以大大降低
 * 内存的耗费，特别使用于大数据量的文件。
 *
 */
public class Excel2007Reader extends DefaultHandler implements IExcelReader {
	//共享字符串表
	private SharedStringsTable sst;
	//上一次的内容
	private String lastContents;
	private boolean nextIsString;

	private int sheetIndex = -1;
	private List<String> rowlist = new ArrayList<String>();
	//当前行
	private int curRow = 0;
	//当前列
	private int curCol = 0;
	private StringBuffer value;
	
	private boolean vIsOpen;
	
	private StylesTable stylesTable;
	
	private ReadOnlySharedStringsTable sharedStringsTable;
	
	/**
	 * 下一个数据类型
	 */
	private XssfDataType nextDataType;
	
	private short formatIndex;
	private String formatString;
	
	private boolean isTElement;
	
	/**
	 * 行读取接口
	 */
	private IRowReader rowReader;
	
	/**
	 * sheet名称
	 */
	private String sheetName;
	
	/**
	 * 文件夹路径
	 */
	private String dirPath;
	
	/**
	 * excel文件名称
	 */
	private String fileName;

	/**
	 * 遍历工作簿中所有的电子表格
	 * @param dirPath
	 * @param fileName
	 * @param rowReader
	 * @throws Exception
     */
	public void readExcel(String dirPath, String fileName, IRowReader rowReader) throws Exception {
		this.dirPath = dirPath;
		this.fileName = dirPath + "/" + fileName;
		this.rowReader = rowReader;
		this.value = new StringBuffer();
		
		OPCPackage pkg = OPCPackage.open(this.fileName);
		XSSFReader xssfReader = new XSSFReader(pkg);
		SAXParserFactory saxFactory = SAXParserFactory.newInstance();
		SAXParser saxParser = saxFactory.newSAXParser();
		XMLReader sheetParser = saxParser.getXMLReader();
		XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		stylesTable = xssfReader.getStylesTable();
		sharedStringsTable = new ReadOnlySharedStringsTable(pkg);
		while (iter.hasNext()) {
			curRow = 0;
			sheetIndex++;
			InputStream sheet = iter.next();
			sheetName = iter.getSheetName();
			InputSource sheetSource = new InputSource(sheet);
			sheetParser.setContentHandler(this);
			sheetParser.parse(sheetSource);
			sheet.close();
		}
		rowReader.close();
		System.out.println("读取【" + fileName + "】完成！");
	}

	public void startElement(String uri, String localName, String name,
			Attributes attributes) throws SAXException {
		
		if ("inlineStr".equals(name) || "v".equals(name)) {
			vIsOpen = true;
			// Clear contents cache
			value.setLength(0);
		}
		// c => cell
		else if ("c".equals(name)) {
			// Get the cell reference
			String r = attributes.getValue("r");
			for (int c = 0; c < r.length(); ++c) {
				if (Character.isDigit(r.charAt(c))) {
					break;
				}
			}
//			thisColumn = nameToColumn(r.substring(0, firstDigit));

			// Set up defaults.
			this.nextDataType = XssfDataType.NUMBER;
			this.formatIndex = -1;
			this.formatString = null;
			String cellType = attributes.getValue("t");
			String cellStyleStr = attributes.getValue("s");
			if ("b".equals(cellType))
				nextDataType = XssfDataType.BOOL;
			else if ("e".equals(cellType))
				nextDataType = XssfDataType.ERROR;
			else if ("inlineStr".equals(cellType))
				nextDataType = XssfDataType.INLINESTR;
			else if ("s".equals(cellType))
				nextDataType = XssfDataType.SSTINDEX;
			else if ("str".equals(cellType))
				nextDataType = XssfDataType.FORMULA;
			else if (cellStyleStr != null) {
				// It's a number, but almost certainly one
				// with a special style or format
				int styleIndex = Integer.parseInt(cellStyleStr);
				XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
				this.formatIndex = style.getDataFormat();
				this.formatString = style.getDataFormatString();
				if (this.formatString == null)
					this.formatString = BuiltinFormats
							.getBuiltinFormat(this.formatIndex);
			}
		}
	}

	public void endElement(String uri, String localName, String name)
			throws SAXException {
		String thisStr = null;
		// 根据SST的索引值的到单元格的真正要存储的字符串
		// 这时characters()方法可能会被调用多次
		if (nextIsString) {
			try {
				int idx = Integer.parseInt(lastContents);
				lastContents = new XSSFRichTextString(sst.getEntryAt(idx))
						.toString();
			} catch (Exception e) {

			}
		} 
		//t元素也包含字符串
		if(isTElement){
			String value = lastContents.trim();
			rowlist.add(curCol, value);
			curCol++;
			isTElement = false;
			// v => 单元格的值，如果单元格是字符串则v标签的值为该字符串在SST中的索引
			// 将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
		} else if ("v".equals(name)) {
			switch (nextDataType) {
			case BOOL:
				char first = value.charAt(0);
				thisStr = first == '0' ? "FALSE" : "TRUE";
				break;
			case ERROR:
				thisStr = "\"ERROR:" + value.toString();
				break;
			case FORMULA:
				// A formula could result in a string value,
				// so always add double-quote characters.
				thisStr = value.toString();
				break;
			case INLINESTR:
				XSSFRichTextString rtsi = new XSSFRichTextString(
						value.toString());
				thisStr = rtsi.toString();
				break;
			case SSTINDEX:
				String sstIndex = value.toString();
				try {
					int idx = Integer.parseInt(sstIndex);
					XSSFRichTextString rtss = new XSSFRichTextString(
							sharedStringsTable.getEntryAt(idx));
					thisStr = rtss.toString();
				} catch (NumberFormatException ex) {
//					output.println("Failed to parse SST index '" + sstIndex
//							+ "': " + ex.toString());
					ex.printStackTrace();
				}
				break;
			case NUMBER:
				String n = value.toString();
//				// 判断是否是日期格式
//				if (HSSFDateUtil.isADateFormat(this.formatIndex, n)) {
//					Double d = Double.parseDouble(n);
//					Date date=HSSFDateUtil.getJavaDate(d);
//					thisStr=formateDateToString(date);
//				} else if (this.formatString != null)
//					thisStr = formatter.formatRawCellContents(
//							Double.parseDouble(n), this.formatIndex,
//							this.formatString);
//				else
					thisStr = n;
				break;

			default:
				thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
				break;
			}
			rowlist.add(curCol, thisStr);
			curCol++;
		}else {
			//如果标签名称为 row ，这说明已到行尾，调用 optRows() 方法
			if (name.equals("row")) {
				try {
					rowReader.getRows(sheetIndex, curRow, sheetName, dirPath, rowlist);
				} catch (Exception e) {
					e.printStackTrace();
				}
				rowlist.clear();
				curRow++;
				curCol = 0;
			}
		}
		
	}

	public void characters(char[] ch, int start, int length)
			throws SAXException {
		//得到单元格内容的值
		if (vIsOpen)
			value.append(ch, start, length);
	}
}

enum XssfDataType {
	BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER,
}