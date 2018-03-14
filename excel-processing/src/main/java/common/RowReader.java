package common;

import org.apache.commons.lang3.StringUtils;

import java.util.ArrayList;
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
public class RowReader implements IRowReader {
    private List<Object> objects = new ArrayList<>();

    public List<Object> getObjects() {
        return objects;
    }

    public void setObjects(List<Object> objects) {
        this.objects = objects;
    }

    /*
             * 业务逻辑实现方法
             *
             * @see com.eprosun.util.excel.IRowReader#getRows(int, int, java.util.List)
             */
    @Override
    public void getRows(int sheetIndex, int curRow, List<String> rowlist) {
        // TODO Auto-generated method stub
        Object object = new Object();
        if (curRow == 0) {
            return;
        }
        if (rowlist.size() == 5) {
            List<String> list = rowlist.subList(0, 2);
            list.add("");
            list.addAll(rowlist.subList(3, 5));
            rowlist = list;
        }
        String value = "";
        for (int i = 1; i < rowlist.size(); i++) {
            value = rowlist.get(i);
            if (StringUtils.isBlank(value) || " ".equals(value)) {
                value = "0";
            }
            /*switch (i) {
                case 1:
                    youzanSkuTemp.setYouzanSku(Integer.parseInt(value));
                    break;
                case 2:
                    youzanSkuTemp.setXfxbSku(Integer.parseInt(value));
                    break;
                case 3:
                    youzanSkuTemp.setXfxbSkuUp(Integer.parseInt(value));
                    break;
                case 4:
                    youzanSkuTemp.setPropertiesName(value);
                    break;

                case 5:
                    youzanSkuTemp.setTitle(value);
                    break;
            }*/
        }
        objects.add(object);
        return;
    }
}
