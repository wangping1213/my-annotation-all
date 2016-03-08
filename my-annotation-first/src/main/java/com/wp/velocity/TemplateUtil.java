package com.wp.velocity;

import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.VelocityEngine;
import org.apache.velocity.runtime.RuntimeConstants;
import org.apache.velocity.runtime.resource.loader.ClasspathResourceLoader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.StringWriter;
import java.util.Map;


/**
 * 模板解析工具
 * @author <a href="">wangping</a>
 * @version 1.0
 * @since 2016/3/3 9:01
 */
public class TemplateUtil {

    private TemplateUtil() {}

    private Logger logger = LoggerFactory.getLogger(this.getClass());

    private static VelocityEngine ve;

    private static void init() {
        ve = new VelocityEngine();
        ve.setProperty(RuntimeConstants.RESOURCE_LOADER, "classpath");//取得当前classpath，文件路径
        ve.setProperty("classpath.resource.loader.class", ClasspathResourceLoader.class.getName());
        ve.setProperty("input.encoding", "UTF-8");
        ve.setProperty("output.encoding", "UTF-8");
        try {
            ve.init();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 解析模板文件，返回解析完之后的字符串
     * @param map
     * @param fileName
     */
    public static String parseTemplate(Map<String, Object> map, String fileName) {
        String str = "";
        if (null == ve) init();
        try {
            Template template = ve.getTemplate(fileName);//取得模板文件
            VelocityContext ctx = new VelocityContext();
            //建议以后map统一只用这种方式，效率比使用map.keySet()高很多***
            for (Map.Entry<String, Object> entry : map.entrySet()) {
                ctx.put(entry.getKey(), entry.getValue());
            }
            StringWriter sw = new StringWriter();
            template.merge(ctx, sw);
            str = sw.toString();
//            System.out.println(str);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return str;
    }
}
