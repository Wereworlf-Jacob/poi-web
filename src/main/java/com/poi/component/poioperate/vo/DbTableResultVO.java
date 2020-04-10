package com.poi.component.poioperate.vo;

/**
 * @ClassName DbTableResultVO
 * @Description 数据表整理结果VO
 * @Author Jacob
 * @Version 1.0
 * @since 2020/4/10 13:44
 **/
public class DbTableResultVO {

    //云代账数据库的表名
    private String ydzTableName;

    //代理版数据库的表名
    private String agencyTableName;

    //云代账自增字段当前值
    private String ydzIncr;

    //代理版自增字段当前值
    private String agencyIncr;

    //更新后的自增字段当前值
    private String updateIncr;

    public String getYdzTableName() {
        return ydzTableName;
    }

    public void setYdzTableName(String ydzTableName) {
        this.ydzTableName = ydzTableName;
    }

    public String getAgencyTableName() {
        return agencyTableName;
    }

    public void setAgencyTableName(String agencyTableName) {
        this.agencyTableName = agencyTableName;
    }

    public String getYdzIncr() {
        return ydzIncr;
    }

    public void setYdzIncr(String ydzIncr) {
        this.ydzIncr = ydzIncr;
    }

    public String getAgencyIncr() {
        return agencyIncr;
    }

    public void setAgencyIncr(String agencyIncr) {
        this.agencyIncr = agencyIncr;
    }

    public String getUpdateIncr() {
        return updateIncr;
    }

    public void setUpdateIncr(String updateIncr) {
        this.updateIncr = updateIncr;
    }
}
