package com.wp.mvc.client;

import com.alibaba.fastjson.JSON;
import com.wp.mvc.annotation.MyAnnotation;
import com.wp.mvc.model.MyHello;

/**
 * 测试
 * @author <a href="">wangping</a>
 * @version 1.0
 * @since 2016/2/23 13:37
 */
@MyAnnotation(name="MyClient")
public class MyClient {
    @MyAnnotation(name="hello")
    private MyHello hello;

    @MyAnnotation(name="main")
    public static void main(String[] args) {
//        MyAnnotationUtil.handleMyAnnotation(MyClient.class);
//        System.out.println(List.class.isAssignableFrom(ArrayList.class));

//        Field[] fields = MyClient.class.getDeclaredFields();
//        for (Field field : fields) {
//            if (field.isAnnotationPresent(MyAnnotation.class)) {
//                MyAnnotation a = field.getAnnotation(MyAnnotation.class);
//                System.out.println(a.name());
//            }
//        }
        String str = "{'name':'5'}";
        MyHello hello = JSON.parseObject(str, MyHello.class);
        System.out.println(hello);
    }
}

class Hello {

    private String name;
}
