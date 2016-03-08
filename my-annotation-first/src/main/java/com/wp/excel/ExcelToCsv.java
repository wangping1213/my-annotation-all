package com.wp.excel;

/**
 * 将excel转为csv（其中有多个sheet，并且每个sheet都存在标题行）
 * 使用csv模式来读取excel，防止出现内存溢出的问题
 * @author wangping
 * @time 2015年12月24日 下午9:27:28
 */
public class ExcelToCsv {
	
	public static void main(String[] args) throws Exception {
        ExcelReaderUtil.readExcel("C:/Users/X1C/Desktop", "机场火车站维表.xlsx");
	}
}