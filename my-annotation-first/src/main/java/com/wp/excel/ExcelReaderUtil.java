package com.wp.excel;

/**
 * excel读取工具（可以用于读取2003或2007）
 * @author wangping
 * @time 2015年12月25日 下午1:05:03
 */
public class ExcelReaderUtil {
	
	/**
	 * excel2003扩展名
	 */
	public static final String EXCEL03_EXTENSION = ".xls";
	
	/**
	 * excel2007扩展名
	 */
	public static final String EXCEL07_EXTENSION = ".xlsx";
	
	/**
	 * 读取Excel文件，可能是03也可能是07版本
	 * @param fileName 文件名
	 * @throws Exception 异常 
	 */
	public static void readExcel(String dirPath, String fileName) throws Exception{
		IRowReader reader = new RowReader();  
		IExcelReader excelReader = null;
		if (fileName.endsWith(EXCEL03_EXTENSION)) {// 处理excel2003文件
			excelReader = new Excel2003Reader();  
		} else if (fileName.endsWith(EXCEL07_EXTENSION)) {// 处理excel2007文件
			excelReader = new Excel2007Reader();  
		} else {
			throw new  Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
		}
		excelReader.readExcel(dirPath, fileName, reader);
	}
}
