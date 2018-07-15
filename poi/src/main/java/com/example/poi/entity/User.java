package com.example.poi.entity;

import com.example.poi.annotation.ExcelModel;
import com.example.poi.annotation.ImportModel;

import java.io.Serializable;

/**
 * Created by wangliang on 2017/8/31.
 * @author wangliang
 */
public class User implements Serializable, ImportModel {

	@ExcelModel(chineseName = "姓名", index = 0)
    private String name;
    @ExcelModel(chineseName = "年龄", index = 1)
    private String age;
    @ExcelModel(chineseName = "性别", index = 2, isStatusValue = true)
    private Integer sex;
    @ExcelModel(chineseName = "邮箱", index = 3)
    private String email;
    @ExcelModel(chineseName = "手机号", index = 4)
    private String phone;

    public User() { }

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
