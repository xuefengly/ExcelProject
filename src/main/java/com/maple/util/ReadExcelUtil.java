package com.maple.util;

import com.alibaba.excel.EasyExcel;
import com.maple.listener.NoModelListener;
import com.maple.listener.ObjectListener;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 读取工具类
 */
public class ReadExcelUtil {

    /**
     * 返回Map型的实体模型数据
     * @param fileName 文件名
     * @param obj  对象反射类 xxx.Class
     * @param sheetNo sheet编号
     * @param headNum 表头行数，1表示1行
     * @return
     */
    public static Map<String, List<Object>> getMap(String fileName, Class obj, Integer sheetNo, Integer headNum){
        Map<String, List<Object>> map = new HashMap<>();
        //new监听器
        ObjectListener objectListener = new ObjectListener();
        //读取操作
        EasyExcel.read(fileName, obj, objectListener).sheet(sheetNo).headRowNumber(headNum).doRead();
        //获取读取的数据
        List<Object> objList = objectListener.getObjectList();
        //获取sheet名
        String objSheetName = objectListener.getSheetName();
        map.put(objSheetName,objList);
        return map;
    }

    //返回List型的实体模型数据
    public static List<Object> getList(String fileName,Class obj,Integer sheetNo,Integer headNum){
        ObjectListener objectListener = new ObjectListener();
        EasyExcel.read(fileName, obj, objectListener).sheet(sheetNo).headRowNumber(headNum).doRead();
        List<Object> objList = objectListener.getObjectList();
        return objList;
    }

    //返回有实体模型的表头
    public static List<Object> getHeadList(String fileName,Class obj,Integer sheetNo,Integer headNum){
        ObjectListener objectListener = new ObjectListener();
        EasyExcel.read(fileName, obj, objectListener).sheet(sheetNo).headRowNumber(headNum).doRead();
        List<Object> objList = objectListener.getHeadList();
        return objList;
    }


    //返回Map型的没有实体模型数据
    public static Map<String,List<Map<Integer, String>>> getMapNoModel(String fileName,Integer sheetNo,Integer headNum){
        Map<String,List<Map<Integer, String>>> map = new HashMap<>();
        NoModelListener noModelListener = new NoModelListener();
        EasyExcel.read(fileName, noModelListener).sheet(sheetNo).headRowNumber(headNum).doRead();
        List<Map<Integer, String>> lists = noModelListener.getLists();
        String sheetName = noModelListener.getSheetName();
        map.put(sheetName,lists);
        return map;
    }

    //返回List型的没有实体模型数据
    public static List<Map<Integer, String>> getListNoModel(String fileName,Integer sheetNo,Integer headNum){
        NoModelListener noModelListener = new NoModelListener();
        EasyExcel.read(fileName, noModelListener).sheet(sheetNo).headRowNumber(headNum).doRead();
        List<Map<Integer, String>> lists = noModelListener.getLists();
        return lists;
    }

    //返回List型的没有实体模型数据(仅有数据)
    public static List<List<Object>> getOnlyListNoModel(String fileName,Integer sheetNo,Integer headNum){
        NoModelListener noModelListener = new NoModelListener();
        EasyExcel.read(fileName, noModelListener).sheet(sheetNo).headRowNumber(headNum).doRead();
        List<Map<Integer, String>> lists = noModelListener.getLists();
        List<List<Object>> allData = new ArrayList<>();
        for (int i = 0; i < lists.size(); i++) {
            Map<Integer, String> map = lists.get(i);
            List<Object> rowData = new ArrayList<>();
            for (Map.Entry<Integer, String> entry : map.entrySet()) {
                rowData.add(entry.getValue());
            }
            allData.add(rowData);
        }
        return allData;
    }

    //返回没有实体模型的表头
    public static List<Map<Integer,String>> getHeadListNoModel(String fileName,Integer sheetNo,Integer headNum){
        NoModelListener noModelListener = new NoModelListener();
        EasyExcel.read(fileName, noModelListener).sheet(sheetNo).headRowNumber(headNum).doRead();
        List<Map<Integer, String>> lists = noModelListener.getHeadList();
        return lists;
    }

    /**
     * 返回没有实体模型的具体行的表头
     * @param fileName : 文件名
     * @param sheetNo : sheet编号
     * @param headNum : 表头行数，1表示1行
     * @param rowNum : 行号，索引从0开始，即0为第一行
     * @return java.util.List<java.lang.String>
     * @throws
     */
    public static List<String> getHeadListNoModel(String fileName,Integer sheetNo,Integer headNum,Integer rowNum){
        NoModelListener noModelListener = new NoModelListener();
        EasyExcel.read(fileName, noModelListener).sheet(sheetNo).headRowNumber(headNum).doRead();
        List<Map<Integer, String>> headList = noModelListener.getHeadList();
        List<String> oneHead = new ArrayList<>();
        for (int i = 0; i < headList.size(); i++) {
            if(i==rowNum){
                Map<Integer, String> map = headList.get(i);
                for (Map.Entry<Integer, String> entry : map.entrySet()) {
                    oneHead.add(entry.getValue());
                }

            }
        }
        return oneHead;
    }
}
