package com.wp.excel;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 抽象Excel2003读取器，通过实现HSSFListener监听器，采用事件驱动模式解析excel2003
 * 中的内容，遇到特定事件才会触发，大大减少了内存的使用。
 * @author wangping
 * @time 2015年12月29日 上午11:08:12
 */
public  class Excel2003Reader implements HSSFListener, IExcelReader {
	private POIFSFileSystem fs;
	private String dirPath;

	/** Should we output the formula, or the value it has? */
	private boolean outputFormulaValues = true;

	/** For parsing Formulas */
	private SheetRecordCollectingListener workbookBuildingListener;
	//excel2003工作薄
	private HSSFWorkbook stubWorkbook;

	// Records we pick up as we process
	private SSTRecord sstRecord;
	private FormatTrackingHSSFListener formatListener;

	//表索引
	private int sheetIndex = -1;
	private BoundSheetRecord[] orderedBSRs;
	
	@SuppressWarnings("rawtypes")
	private List boundSheetRecords = new ArrayList();

	// For handling formulas with string results
	private int nextRow;
	private int nextColumn;
	
	/**
	 * 
	 */
	private boolean outputNextStringRecord;
	
	/**
	 * 当前行序号（从0开始）
	 */
	private int curRow = 0;
	
	/**
	 * 存储行记录的容器（存储行数据列表）
	 */
	private List<String> rowlist = new ArrayList<String>();;

	/**
	 * sheet名称
	 */
	private String sheetName;
	
	/**
	 * 行读取对象
	 */
	private IRowReader rowReader;
	
	/**
	 * 2003的excel文件名（无路径）
	 */
	private String fileName;
	
	
	/**
	 * 遍历excel下所有的sheet
	 * @throws IOException
	 */
	public void readExcel(String dirPath, String fileName, IRowReader rowReader) throws Exception {
		this.dirPath = dirPath;
		this.fileName = dirPath + "/" + fileName;
		this.rowReader = rowReader;

		this.fs = new POIFSFileSystem(new FileInputStream(this.fileName));
		MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
		formatListener = new FormatTrackingHSSFListener(listener);
		HSSFEventFactory factory = new HSSFEventFactory();
		HSSFRequest request = new HSSFRequest();
		if (outputFormulaValues) {
			request.addListenerForAllRecords(formatListener);
		} else {
			workbookBuildingListener = new SheetRecordCollectingListener(formatListener);
			request.addListenerForAllRecords(workbookBuildingListener);
		}
		factory.processWorkbookEvents(request, fs);
		rowReader.close();

		System.out.println("读取【" + fileName + "】完成！");
	}
	
	/**
	 * HSSFListener 监听方法，处理 Record
	 */
	@SuppressWarnings("unchecked")
	public void processRecord(Record record) {
		int thisColumn = -1;
		String thisStr = null;
		String value = null;
		switch (record.getSid()) {
			case BoundSheetRecord.sid:
				boundSheetRecords.add(record);
				break;
			case BOFRecord.sid:
				BOFRecord br = (BOFRecord) record;
				if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
					// 如果有需要，则建立子工作薄
					if (workbookBuildingListener != null && stubWorkbook == null) {
						stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
					}
					
					sheetIndex++;
					if (orderedBSRs == null) {
						orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
					}
					sheetName = orderedBSRs[sheetIndex].getSheetname();
				}
				break;
	
			case SSTRecord.sid:
				sstRecord = (SSTRecord) record;
				break;
	
			case BlankRecord.sid:
				BlankRecord brec = (BlankRecord) record;
				curRow = brec.getRow();
				thisColumn = brec.getColumn();
				thisStr = "";
				rowlist.add(thisColumn, thisStr);
				break;
			case BoolErrRecord.sid: //单元格为布尔类型
				BoolErrRecord berec = (BoolErrRecord) record;
				curRow = berec.getRow();
				thisColumn = berec.getColumn();
				thisStr = berec.getBooleanValue()+"";
				rowlist.add(thisColumn, thisStr);
				break;
	
			case FormulaRecord.sid: //单元格为公式类型
				FormulaRecord frec = (FormulaRecord) record;
				curRow = frec.getRow();
				thisColumn = frec.getColumn();
				if (outputFormulaValues) {
					if (Double.isNaN(frec.getValue())) {
						// Formula result is a string
						// This is stored in the next record
						outputNextStringRecord = true;
						nextRow = frec.getRow();
						nextColumn = frec.getColumn();
					} else {
						thisStr = formatListener.formatNumberDateCell(frec);
					}
				} else {
					thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook,
							frec.getParsedExpression()) + '"';
				}
				rowlist.add(thisColumn,thisStr);
				break;
			case StringRecord.sid://单元格中公式的字符串
				if (outputNextStringRecord) {
					// String for formula
					StringRecord srec = (StringRecord) record;
					thisStr = srec.getString();
					curRow = nextRow;
					thisColumn = nextColumn;
					outputNextStringRecord = false;
				}
				break;
			case LabelRecord.sid:
				LabelRecord lrec = (LabelRecord) record;
				curRow = lrec.getRow();
				thisColumn = lrec.getColumn();
				value = lrec.getValue().trim();
				value = value.equals("")?" ":value;
				this.rowlist.add(thisColumn, value);
				break;
			case LabelSSTRecord.sid:  //单元格为字符串类型
				LabelSSTRecord lsrec = (LabelSSTRecord) record;
				curRow = lsrec.getRow();
				thisColumn = lsrec.getColumn();
				if (sstRecord == null) {
					rowlist.add(thisColumn, " ");
				} else {
					value =  sstRecord
					.getString(lsrec.getSSTIndex()).toString().trim();
					value = value.equals("")?" ":value;
					rowlist.add(thisColumn,value);
				}
				break;
			case NumberRecord.sid:  //单元格为数字类型
				NumberRecord numrec = (NumberRecord) record;
				curRow = numrec.getRow();
				thisColumn = numrec.getColumn();
				value = formatListener.formatNumberDateCell(numrec).trim();
				value = value.equals("")?" ":value;
				// 向容器加入列值
				rowlist.add(thisColumn, value);
				break;
			default:
				break;
		}

		// 空值的操作
		if (record instanceof MissingCellDummyRecord) {
			MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
			curRow = mc.getRow();
			thisColumn = mc.getColumn();
			rowlist.add(thisColumn," ");
		}

		// 行结束时的操作
		if (record instanceof LastCellOfRowDummyRecord) {
				// 每行结束时， 调用getRows() 方法
			try {
				rowReader.getRows(sheetIndex, curRow, sheetName, dirPath, rowlist);
			} catch (Exception e) {
				e.printStackTrace();
			}
			
			// 清空容器
			rowlist.clear();
		}
	}
	
}