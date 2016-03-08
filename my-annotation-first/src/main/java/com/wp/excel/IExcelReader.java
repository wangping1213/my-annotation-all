package com.wp.excel;

/**
 * 读取excel的接口
 * @author wangping
 * @time 2015年12月25日 下午12:48:53
 */
public interface IExcelReader {
	
	/**
	 * 遍历工作簿中所有的电子表格
	 * @param filename
	 * @throws Exception
	 */
	public void readExcel(String dirPath, String fileName, IRowReader rowReader) throws Exception;
}
