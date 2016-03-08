package com.wp.velocity;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 根据模板和数据导出对应的excel
 * @author wangping
 * @version 1.0
 * @since 2016/3/3 10:26
 */
public class TableToExcel {

    public static void main(String[] args) throws Exception {

        String fileName = "templates/table.vm";
        Map<String, Object> map = new HashMap<String, Object>();
        //类Person必须是public的，也就是说必须是一个单独的类文件***
        List<Person> temp = new ArrayList<Person>();
        temp.add(new Person("111", 1, "no1"));
        temp.add(new Person("222", 2, "no2"));
        map.put("list", temp);

        String result = TemplateUtil.parseTemplate(map, fileName);
        System.out.println(result);
        try {
            String htmlStr = result;
            String sheetName = "222";
            //生成Excel工作薄对象
            HSSFWorkbook wb = ParseHtmlToXls.parseHtmlToXlsForCommon(htmlStr, sheetName);
            OutputStream outputStream = new FileOutputStream("C:\\Users\\X1C\\Desktop\\222.xls");
            wb.write(outputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
