package com.wp.mvc.model;

import com.alibaba.fastjson.annotation.JSONField;

/**
 * 测试
 * @author <a href="">wangping</a>
 * @version 1.0
 * @since 2016/2/23 13:01
 */
public class MyHello {
    @JSONField(name = "name")
    private String name = "";

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
