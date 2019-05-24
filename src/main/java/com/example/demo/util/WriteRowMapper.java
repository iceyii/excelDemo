package com.example.demo.util;

import java.util.List;

public interface WriteRowMapper {
    /**.
     * 方法：在导出excel时，用于讲导出的数据组成list，调用者实现
     * @param param 需要导出到excel中的对象
     * @return
     */
    List<String> handleData(Object param);
    //实现类示例
    //public class ExcelMonthRowMapper implements WriteRowMapper {
    //    public static final String[] TITLES = {"指标编号", "指标名称", "一级分类", "二级分类",
    //            "指标单位", "指标值", "指标日期", "当前预估排名", "状态", "异常原因", "负责人"};
    //
    //    @Override
    //    public List<String> handleData(Object param) {
    //        List<String> value = new ArrayList<>(TITLES.length);
    //        KpiInfoVO kpiInfo = (KpiInfoVO) param;
    //        value.add(kpiInfo.getKpiId());
    //        value.add(kpiInfo.getKpiName3());
    //        value.add(kpiInfo.getKpiName1());
    //        value.add(kpiInfo.getKpiName2());
    //        if ("0".equals(kpiInfo.getKpiType())) {
    //            value.add("%");
    //        } else {
    //            value.add("秒");
    //        }
    //        value.add(String.valueOf(kpiInfo.getKpiValue()));
    //        value.add(kpiInfo.getKpiDate());
    //        value.add("--");
    //        value.add("--");
    //        value.add(kpiInfo.getExceptionReason());
    //        value.add(kpiInfo.getBark1());
    //        return value;
    //    }


    //调用方法
    //    ExcelUtil excelUtil = new ExcelUtil(response);
    //    excelUtil.exportExcel("月异常信息", "异常指标月报", kpiDate, ExcelMonthRowMapper.TITLES, list, new ExcelMonthRowMapper(), fileName);
}
