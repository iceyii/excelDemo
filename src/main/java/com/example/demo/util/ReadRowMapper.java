package com.example.demo.util;

import org.apache.poi.ss.usermodel.Row;

import java.util.Map;

public interface ReadRowMapper<T> {
    /**.
     * 方法：在导入excel时，用于自定义处理excel中的每行数据,调用者实现
     * @param row excel中的行数据
     */
    T rowMap(Row row, Map<String, Object> map);

}
//调用方法
//    List<KpiConfig> list = excelUtil.readExcel(file, (row, map) -> {
//        KpiConfig kpiConfig = new KpiConfig();
//        //id为空就排除此对象
//        if (StringUtils.isBlank((String) map.get("指标编号"))) {
//            return null;
//        }
//        kpiConfig.setKpiName1((String) map.get("一级分类"));
//        kpiConfig.setKpiName2((String) map.get("二级分类"));
//        kpiConfig.setOwner((String) map.get("负责人"));
//        kpiConfig.setDescribe((String) map.get("指标说明"));
//        return kpiConfig;
//    });