package com.example.poi.entity;

/**
 * Created by wangliang on 2017/8/31.
 */
public class User {

    private String name;
    private String age;
    private Integer sex;
    private String email;
    private String phone;

    public User() {
    }

    public User(String name, String age, Integer sex, String email, String phone) {
        this.name = name;
        this.age = age;
        this.sex = sex;
        this.email = email;
        this.phone = phone;
    }

    public String getName() {
        return name;
    }

    public String getAge() {
        return age;
    }

    public Integer getSex() {
        return sex;
    }

    public String getEmail() {
        return email;
    }

    public String getPhone() {
        return phone;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setAge(String age) {
        this.age = age;
    }

    public void setSex(Integer sex) {
        this.sex = sex;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }
}
