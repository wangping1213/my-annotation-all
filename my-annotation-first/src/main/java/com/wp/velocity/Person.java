package com.wp.velocity;

/**
 * Created by X1C on 2016/3/3.
 */
public class Person {

    private String name;
    private int age;
    private String no;

    public Person(String name, int age, String no) {
        this.name = name;
        this.age = age;
        this.no = no;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getNo() {
        return no;
    }

    public void setNo(String no) {
        this.no = no;
    }
}
