package com.wp.mvc.web;

import java.io.Serializable;

/**
 * 人员
 * @author wangping
 * @version 1.0
 * @since 2016/3/17 15:46
 */
public class Person implements Serializable {

    private String name;
    private String nick;
    private Integer age;

    public Person(String name, String nick, Integer age) {
        this.name = name;
        this.nick = nick;
        this.age = age;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getNick() {
        return nick;
    }

    public void setNick(String nick) {
        this.nick = nick;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    @Override
    public String toString() {
        return "Person{" +
                "name='" + name + '\'' +
                ", nick='" + nick + '\'' +
                ", age=" + age +
                '}';
    }
}
