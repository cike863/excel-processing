package common;

/**
 * @version: v1.0
 * @author: zhulong
 * project: order
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/3/1 0001 下午 4:44
 * modifyTime:
 * modifyBy:
 */
public class ExcelReaderUtil {
    //excel2003扩展名
    public static final String EXCEL03_EXTENSION = ".xls";
    //excel2007扩展名
    public static final String EXCEL07_EXTENSION = ".xlsx";

    /**
     * 读取Excel文件，可能是03也可能是07版本
     *
     * @param reader
     * @param fileName
     * @param fileName
     * @throws Exception
     */
    public static void readExcel(IRowReader reader, String fileName) throws Exception {
        // 处理excel2003文件
        if (fileName.endsWith(EXCEL03_EXTENSION)) {
            Excel2003Reader exce2003 = new Excel2003Reader();
            exce2003.setRowReader(reader);
            exce2003.process(fileName);
            // 处理excel2007文件
        } else if (fileName.endsWith(EXCEL07_EXTENSION)) {
            Excel2007Reader exce2007 = new Excel2007Reader();
            exce2007.setRowReader(reader);
            exce2007.process(fileName);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
    }
}
