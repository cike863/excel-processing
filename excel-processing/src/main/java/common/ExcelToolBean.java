package common;

import java.io.Serializable;

/**
 * @version: v1.0
 * @author: zhulong
 * project: order
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/3/1 0001 下午 5:09
 * modifyTime:
 * modifyBy:
 */
public class ExcelToolBean implements Serializable {

    private static final long serialVersionUID = 7798556000531943254L;
    private String operationName;
    private String propertyName;
    private String propertyChinseName;
    private int size;
    private String tableName;
    private String isShow;
    private int location;

    public String getOperationName() {
        return operationName;
    }

    public void setOperationName(String operationName) {
        this.operationName = operationName == null ? null : operationName.trim();
    }

    public String getPropertyName() {
        return propertyName;
    }

    public void setPropertyName(String propertyName) {
        this.propertyName = propertyName == null ? null : propertyName.trim();
    }

    public String getPropertyChinseName() {
        return propertyChinseName;
    }

    public void setPropertyChinseName(String propertyChinseName) {
        this.propertyChinseName = propertyChinseName == null ? null : propertyChinseName.trim();
    }

    public int getSize() {
        return size;
    }

    public void setSize(int size) {
        this.size = size;
    }

    public String getTableName() {
        return tableName;
    }

    public void setTableName(String tableName) {
        this.tableName = tableName == null ? null : tableName.trim();
    }

    public String getIsShow() {
        return isShow;
    }

    public void setIsShow(String isShow) {
        this.isShow = isShow == null ? null : isShow.trim();
    }

    public int getLocation() {
        return location;
    }

    public void setLocation(int location) {
        this.location = location;
    }
}
