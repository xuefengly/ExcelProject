package com.maple.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * @author ：maple
 * @description：学生类
 * @date ：Created in 2020/11/17 16:41
 */
@Data
public class Student {
    //这个注解用于对应表头，value为表头值，index为列值
    @ExcelProperty(value = "姓名",index = 0)
    private String name;

    @ExcelProperty(value = "年龄",index = 1)
    private int age;

    public Student(String name, int age) {
        this.name = name;
        this.age = age;
    }
}
