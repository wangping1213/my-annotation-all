package com.wp.velocity;

import java.util.*;

/**
 * velocity的测试代码-hello world
 * @author <a href="">wangping</a>
 * @version 1.0
 * @since 2016/3/2 22:31
 */
public class HelloVelocity {
    public static void main(String[] args) throws Exception {

        String fileName = "templates/hellovelocity.vm";
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("name", "velocity");
        map.put("date", (new Date()).toString());
        List temp = new ArrayList();
        temp.add("1");
        temp.add("2");
        map.put("list", temp);

//        String result = TemplateUtil.parseTemplate(map, fileName);
//
//        System.out.println();
    }

}
