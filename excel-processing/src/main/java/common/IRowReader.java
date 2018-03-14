package common;

import java.util.List;

/**
 * @version: v1.0
 * @author: zhulong
 * project: order
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/2/16 0016 下午 3:54
 * modifyTime:
 * modifyBy:
 */
public interface IRowReader {
    /**
     * 业务逻辑实现方法
     *
     * @param sheetIndex
     * @param curRow
     * @param rowlist
     */
    public void getRows(int sheetIndex, int curRow, List<String> rowlist);
}
