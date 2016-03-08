package com.wp.mvc.utils;

import com.wp.mvc.annotation.MyAnnotation;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

/**
 * Created by X1C on 2016/2/23.
 */
public class MyAnnotationUtil {

    /**
     * 根据当前的类对象处理对应的注解MyAnnotation
     * @param clazz
     */
    public static void handleMyAnnotation(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(MyAnnotation.class)) {
                MyAnnotation a = field.getAnnotation(MyAnnotation.class);
                System.out.println(a.name());
            }
        }

        if (clazz.isAnnotationPresent(MyAnnotation.class)) {
            MyAnnotation a = clazz.getAnnotation(MyAnnotation.class);
            System.out.println(a.name());
        }

        Method[] methods = clazz.getDeclaredMethods();
        for (Method method : methods) {
            MyAnnotation a = method.getAnnotation(MyAnnotation.class);
            System.out.println(a.name());
        }
    }
}
