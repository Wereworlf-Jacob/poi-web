package com.poi.component.poioperate.vo;

/**
 * @ClassName DbTableIncrement
 * @Description 数据库表自增数据
 * @Author Jacob
 * @Version 1.0
 * @since 2020/4/10 12:02
 **/
public class DbTableIncrement {

    private String tableName;

    private String increment;

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName;
    }

    public String getIncrement() {
        return increment;
    }

    public void setIncrement(String increment) {
        this.increment = increment;
    }
}
