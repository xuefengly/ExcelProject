package com.maple.execute;

import com.alibaba.excel.ExcelWriter;
import com.maple.entity.Cat;
import com.maple.entity.Student;
import com.maple.util.ReadExcelUtil;
import com.maple.util.WriteExcelUtil;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author ：maple
 * @description：
 * @date ：Created in 2020/11/17 17:03
 */
public class Execute {
    public static void main(String[] args) {

        //读
        /*//有实体
        String fileName = "F:\\sunyard\\testFile\\excel\\stu.xlsx";
        Map<String, List<Object>> excelMap = new HashMap<String, List<Object>>();
        //第一张sheet--Student
        Map<String, List<Object>> m1 = ReadExcelUtil.getMap(fileName, Student.class, 0, 1);
        //第二张sheet--Cat
        Map<String, List<Object>> m2 = ReadExcelUtil.getMap(fileName, Cat.class, 1, 1);
        excelMap.putAll(m1);
        excelMap.putAll(m2);
        System.out.println(excelMap);//{Student=[Student(name=AA, age=11), Student(name=BB, age=22), Student(name=AADD, age=33), Student(name=AACC, age=11), Student(name=AAEE, age=5)], Cat=[Cat(name=aa, age=1), Cat(name=bb, age=2), Cat(name=cc, age=3), Cat(name=dd, age=4), Cat(name=ee, age=1)]}*/

        //无实体
        /*
        String fileName = "F:\\sunyard\\testFile\\excel\\stu.xlsx";
        Map<String, List<Map<Integer, String>>> excelMap = new HashMap<>();
        Map<String, List<Map<Integer, String>>> stuMap = ReadExcelUtil.getMapNoModel(fileName, 0, 1);
        Map<String, List<Map<Integer, String>>> catMap = ReadExcelUtil.getMapNoModel(fileName, 1, 1);
        excelMap.putAll(stuMap);
        excelMap.putAll(catMap);
        System.out.println(excelMap);//{Student=[{0=AA, 1=11}, {0=BB, 1=22}, {0=AADD, 1=33}, {0=AACC, 1=11}, {0=AAEE, 1=5}], Cat=[{0=aa, 1=1}, {0=bb, 1=2}, {0=cc, 1=3}, {0=dd, 1=4}, {0=ee, 1=1}]}
        */


        //写
        //有实体
//        List<Object> stuList = new ArrayList<>();
//        Object s1 = new Student("小明",10);
//        Object s2 = new Student("小红",20);
//        stuList.add(s1);
//        stuList.add(s2);
//        List<Object> catList = new ArrayList<>();
//        Object c1 = new Cat("小明",10);
//        Object c2 = new Cat("小红",20);
//        catList.add(c1);
//        catList.add(c2);
//        String file = "F:/sunyard/testFile/excel/writeTest"+System.currentTimeMillis()+".xlsx";
//        ExcelWriter excelWriter = WriteExcelUtil.getExcelWriter(file);
//        WriteExcelUtil.writeMultipleSheet(excelWriter,stuList,Student.class,0,"Student");
//        WriteExcelUtil.writeMultipleSheet(excelWriter,stuList,Cat.class,1,"Cat");
//        WriteExcelUtil.closeExcelWriter(excelWriter);

        //无实体写
        //此处造数据较麻烦，借无实体都获取数据再写入另一个文件
        //读取
        String fileName = "F:\\sunyard\\testFile\\excel\\stu.xlsx";
        //数据
        List<List<Object>> stuList = ReadExcelUtil.getOnlyListNoModel(fileName, 0, 1);
        //表头
        List<String> headList = ReadExcelUtil.getHeadListNoModel(fileName,  0, 1,0);
        List<List<String>> writeHeadList = new ArrayList<>();
        for (int i = 0; i < headList.size(); i++) {
            List<String> oneHead = new ArrayList<>();
            oneHead.add(headList.get(i));
            writeHeadList.add(oneHead);
        }
        //写入
        String file = "F:/sunyard/testFile/excel/writeTest"+System.currentTimeMillis()+".xlsx";
        ExcelWriter excelWriter = WriteExcelUtil.getExcelWriter(file);
        WriteExcelUtil.writeMultipleNoModel(excelWriter,stuList,writeHeadList,0,"Student");
        WriteExcelUtil.closeExcelWriter(excelWriter);

    }
}
