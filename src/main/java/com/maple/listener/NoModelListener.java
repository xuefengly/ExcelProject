package com.maple.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @description：无实体模型的监听器，返回List/Map
 * @date ：Created in 2020/10/30 18:52
 */
public class NoModelListener extends AnalysisEventListener<Map<Integer,String>> {
    private final static Logger LOGGER = LoggerFactory.getLogger(NoModelListener.class);

    private static final int BATCH_COUNT = 5;

    //数据存储结构
    private List<Map<Integer,String>> lists = new ArrayList<>();
    //表头存储结构
    List<Map<Integer,String>> headList = new ArrayList<>();
    private String sheetName;

    List<Map<Integer,String>> datas = new ArrayList<Map<Integer,String>>();

    @Override
    public void invoke(Map<Integer,String> o, AnalysisContext analysisContext) {
//        LOGGER.info("解析到一条数据:{}", JSON.toJSONString(o));
        datas.add(o);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (datas.size() >= BATCH_COUNT) {
            saveData();
            // 存储完成清理 list
            datas.clear();
        }
    }

    //获取表头
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        headList.add(headMap);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        //获取sheetName
        sheetName = context.readSheetHolder().getSheetName();
        saveData();
//        LOGGER.info("所有数据解析完成！");
    }

    /**
     * 入库
     */
    private void saveData() {
//        LOGGER.info("{}条数据，开始存储数据库！", datas.size());
        lists.addAll(datas);

    }

    public List<Map<Integer,String>> getDatas() {
        return datas;
    }

    public void setDatas(List<Map<Integer,String>> datas) {
        this.datas = datas;
    }

    public List<Map<Integer, String>> getLists() {
        return lists;
    }

    public void setLists(List<Map<Integer, String>> lists) {
        this.lists = lists;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<Map<Integer,String>> getHeadList() {
        return headList;
    }

    public void setHeadList(List<Map<Integer,String>> headList) {
        this.headList = headList;
    }
}
