package writer;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @version: v1.0
 * @author: Administrator
 * @Description: project: sysmanager
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/12/4 0004 上午 11:26
 * modifyTime:
 * modifyBy:
 */
public class XSSFCellStyleMap {

    public static final String CELL_STRING = "cell_string";
    public static final String CELL_LONG = "cell_long";
    public static final String CELL_DOUBLE = "cell_double";
    public static final String CELL_DATE = "cell_date";
    public static final String COL_TITLE = "col_title";
    public static final String SHEET_TITLE = "sheet_title";
    public static final Map<String, String> XSSFCellStyleMap = new LinkedHashMap<>();

    {
        XSSFCellStyleMap.put(CELL_STRING, CELL_STRING);
    }
}
