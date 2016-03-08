package com.wp.excel;

import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.util.List;

/**
 * 读取的每一个
 * @author wangping
 * @time 2015年12月25日 上午10:50:26
 */
public class RowReader implements IRowReader{  
	
	/**
	 * 写文件流
	 */
	private OutputStreamWriter write;
	
	/**
	 * 每次提交数量
	 */
	private int commitRows = 5000;
	  
	  
	/**
	 * 
	 * @author wangping
	 * @time 2015年12月25日 上午10:37:23
	 * @param sheetIndex
	 * @param curRow
	 * @param sheetName
	 * @param dirPath
	 * @param rowlist
	 * @throws Exception 
	 */
    public void getRows(int sheetIndex, int curRow, String sheetName, String dirPath, List<String> rowlist) throws Exception {  
    	if (curRow == 0) {
    		String destPath = dirPath + "/" + sheetName + ".csv";
    		if (sheetIndex > 0) {
    			write.flush();
    			write.close();
    		}
    		write = new OutputStreamWriter(new FileOutputStream(destPath), "UTF-8"); 
    		return;
    	}
        
        StringBuffer append = new StringBuffer();
		for (String str : rowlist) {
			if (append.toString().length() != 0) {
				append.append("\t");
			}
			append.append((null == str) ? "" : str);
		}
//		append.append("\n");
		System.out.println(append.toString());
		try {
			write.append(append.toString());
			if (curRow % commitRows == 0) {
				write.flush();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
    }
    

    /**
	 * 关闭
	 * @author wangping
	 * @time 2015年12月25日 上午10:52:19
	 */
	public void close() {
		try {
			if (null != write) {
				write.flush();
				write.close();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}  
