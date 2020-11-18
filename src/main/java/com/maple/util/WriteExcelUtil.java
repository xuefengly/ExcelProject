package com.maple.util;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.write.metadata.WriteSheet;

import java.util.List;

/**
 * 写入工具类
 */
public class WriteExcelUtil {

    // 有实体单独写一个sheet
    public static void writeSheet(String fileName,List<Object> list,Class obj,Integer sheetNo,String sheetName){
        ExcelWriter excelWriter = EasyExcel.write(fileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet(sheetNo, sheetName).head(obj).build();
        excelWriter.write(list, writeSheet);
        //千万别忘记finish 会帮忙关闭流
        excelWriter.finish();
    }

    //获取流
    public static ExcelWriter getExcelWriter(String fileName){
        ExcelWriter excelWriter = EasyExcel.write(fileName).build();
        return excelWriter;
    }
    //关闭流
    public static void closeExcelWriter(ExcelWriter excelWriter){
        //千万别忘记finish 会帮忙关闭流
        excelWriter.finish();
    }

    /**
     * 有实体多个sheet一起写，要使用同一个流
     * @param excelWriter : 写入流
     * @param list : 数据
     * @param obj : 对应实体类
     * @param sheetNo : sheet编号，0开始
     * @param sheetName : sheetName
     * @return void
     * @throws
     */
    public static void writeMultipleSheet(ExcelWriter excelWriter,List<Object> list,Class obj,Integer sheetNo,String sheetName){
        WriteSheet writeSheet = EasyExcel.writerSheet(sheetNo, sheetName).head(obj).build();
        excelWriter.write(list, writeSheet);
    }

    //
    /**
     * 无实体多个sheet一起写，要使用同一个流
     * @param excelWriter : 写入流
     * @param list : 数据,List<List<Object>>格式,每一个List<Object>代表一行数据
     * @param headList : 表头，List<List<String>,每一个List<String>代表一列的表头
     * @param sheetNo : sheet编号
     * @param sheetName : sheetName
     * @return void
     * @throws
     */
    public static void writeMultipleNoModel(ExcelWriter excelWriter,List<List<Object>> list,List<List<String>> headList,Integer sheetNo,String sheetName){
        WriteSheet writeSheet = EasyExcel.writerSheet(sheetNo, sheetName).head(headList).build();
        excelWriter.write(list, writeSheet);
    }

}
