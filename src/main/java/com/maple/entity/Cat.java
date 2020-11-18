package com.maple.entity;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * @author ：maple
 * @description：
 * @date ：Created in 2020/11/17 16:45
 */
@Data
public class Cat {
    @ExcelProperty(value = "昵称",index = 0)
    private String name;

    @ExcelProperty(value = "年龄",index = 1)
    private int age;

    public Cat(String name, int age) {
        this.name = name;
        this.age = age;
    }
}
