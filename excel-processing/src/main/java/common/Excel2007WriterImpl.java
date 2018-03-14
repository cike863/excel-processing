package common;

/**
 * @version: v1.0
 * @author: zhulong
 * project: order
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/3/1 0001 下午 4:48
 * modifyTime:
 * modifyBy:
 */
public class Excel2007WriterImpl extends AbstractExcel2007Writer {


    /**
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub
        System.out.println("............................");
        long start = System.currentTimeMillis();
        //构建excel2007写入器
        AbstractExcel2007Writer exce2007Writer = new Excel2007WriterImpl();
        //调用处理方法
        exce2007Writer.process("D://test125.xlsx");
        long end = System.currentTimeMillis();
        System.out.println("....................." + (end - start) / 1000);
    }


    /*
     * 可根据需求重写此方法，对于单元格的小数或者日期格式，会出现精度问题或者日期格式转化问题，建议使用字符串插入方法
     * @see com.excel.ver2.AbstractExcel2007Writer#generate()
     */
    @Override
    public void generate() throws Exception {
        //电子表格开始
        beginSheet();
        for (int rownum = 0; rownum < 150000; rownum++) {
            //插入新行
            insertRow(rownum);
            //建立新单元格,索引值从0开始,表示第一列
            createCell(0, "中国" + rownum + "!");
            createCell(1, 34343.123456789);
            createCell(2, "23.67%");
            createCell(3, "12:12:23");
            createCell(4, "2010-10-11 12:12:23");
            createCell(5, "true");
            createCell(6, "false");
            createCell(7, "false");
            createCell(8, "2010-10-11 12:12:23");
            createCell(9, "true");
            createCell(10, "false");
            createCell(11, "false");
            createCell(12, "2010-10-11 12:12:23");
            createCell(13, "true");
            createCell(14, "false");
            createCell(15, "false");
            //结束行
            endRow();
        }
        //电子表格结束
        endSheet();
    }

    /**
     * 类使用者应该使用此方法进行写操作
     *
     * @param sw
     * @throws Exception
     */
    @Override
    public void generate(SpreadsheetWriter sw) throws Exception {
        //电子表格开始
        sw.beginSheetStyle();
        sw.defaultRowHeight(null);
        sw.beginColsStyle();
        int i = 10;
        for (int rownum = 0; rownum < i; rownum++) {
            sw.defaultColsStyle(rownum+1, (rownum + 1) * 10);
        }
        sw.endColsStyle();
        sw.beginSheetData();
        for (int rownum = 0; rownum < 5; rownum++) {
            //插入新行
            sw.insertStyleRow(rownum, 30 + rownum);
            //建立新单元格,索引值从0开始,表示第一列
            createCell(0, "中国" + rownum + "!");
            createCell(1, 34343.123456789);
            createCell(2, "23.67%");
            createCell(3, "12:12:23");
            createCell(4, "2010-10-11 12:12:23");
            createCell(5, "true");
            createCell(6, "false");
            createCell(7, "false");
            createCell(8, "2010-10-11 12:12:23");
            createCell(9, "true");
            //结束行
            endRow();
        }
        //电子表格结束
        sw.endSheetStyle();
    }

}
