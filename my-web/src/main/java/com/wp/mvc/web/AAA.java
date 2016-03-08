package com.wp.mvc.web;

import com.alibaba.fastjson.JSON;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import java.util.ArrayList;
import java.util.List;

/**
 * 测试
 * @author <a href="">wangping</a>
 * @version 1.0
 * @since 2016/2/28 17:15
 */
@Controller
public class AAA {

    private Logger logger = LoggerFactory.getLogger(this.getClass());

    @RequestMapping("/aaa")
    public String aaa() {
        System.out.println("aaa");
        logger.info("aaa:hahaha");
        return "aaa";
    }

    @RequestMapping("/jsonResult")
    @ResponseBody
    public List<String> jsonResult() {
        System.out.println("jsonResult");
        List<String> list = new ArrayList<String>();
        list.add("111");
        list.add("222");
        logger.info("jsonResult2:{}", JSON.toJSON(list));
        return list;
    }
}
