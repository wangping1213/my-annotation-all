package com.wp.mvc.web;

import com.alibaba.fastjson.JSON;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
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

    @RequestMapping("/hello")
    public String hello(Model model) {
        System.out.println("hello");
        List<Person> temp = new ArrayList<Person>();
        temp.add(new Person("111", "no1", 1));
        temp.add(new Person("222", "no2", 2));
        model.addAttribute("list", temp);
        logger.info("go into hello method!");
        return "hello";
    }

    @RequestMapping("/getPerson")
    @ResponseBody
    public Person getPerson() {
        return new Person("name", "nick", 11);
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
