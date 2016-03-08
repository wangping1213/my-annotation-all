package com.wp.excel;

import java.util.List;

public interface IRowReader {
	
	/**业务逻辑实现方法
	 * @param sheetIndex
	 * @param curRow
	 * @param rowlist
	 */
	public void getRows(int sheetIndex, int curRow, String sheetName, String dirPath, List<String> rowlist) throws Exception;
	
	/**
	 * 关闭资源
	 * @author wangping
	 * @time 2015年12月25日 上午10:52:19
	 */
	public void close(); 
}
