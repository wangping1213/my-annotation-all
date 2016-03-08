package com.wp.excel;

public class ExcelReaderTest {
	public static void main(String[] args) {
		IRowReader iRowReader = SpiLoader.getLoader(IRowReader.class);
		IExcelReader iExcelReader = new Excel2007Reader();

		if (null == iRowReader || null == iExcelReader)
			return;
		try {
			iExcelReader.readExcel("C:/Users/X1C/Desktop", "规则维表0113.xlsx", iRowReader);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}