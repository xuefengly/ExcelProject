package com.maple.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 通用监听器
 */
public class ObjectListener extends AnalysisEventListener<Object> {
    private final static Logger LOGGER = LoggerFactory.getLogger(ObjectListener.class);

    private static final int BATCH_COUNT = 5;

    List<Object> objectList = new ArrayList<>();
    List<Object> headList = new ArrayList<>();
    private String sheetName;

    List<Object> datas = new ArrayList<Object>();

    @Override
    public void invoke(Object o, AnalysisContext analysisContext) {
        LOGGER.info("解析到一条数据:{}", JSON.toJSONString(o));
        //一条数据添加到暂时存储的存储结构中
        datas.add(o);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (datas.size() >= BATCH_COUNT) {
            saveData();
            // 存储完成清理 list
            datas.clear();
        }
    }



    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        //获取sheetName
        sheetName = context.readSheetHolder().getSheetName();
//        LOGGER.info("所有数据解析完成！");
    }

    //获取表头
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        //把表头数据加入到存储结构中
        headList.add(headMap);
    }

    /**
     * 入库
     */
    private void saveData() {
        LOGGER.info("{}条数据，开始存储数据库！", datas.size());
        //添加到返回的存储结构中，也可直接存储到数据库
        objectList.addAll(datas);
    }

    public List<Object> getDatas() {
        return datas;
    }

    public void setDatas(List<Object> datas) {
        this.datas = datas;
    }

    public List<Object> getObjectList() {
        return objectList;
    }

    public void setObjectList(List<Object> objectList) {
        this.objectList = objectList;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<Object> getHeadList() {
        return headList;
    }

    public void setHeadList(List<Object> headList) {
        this.headList = headList;
    }
}
