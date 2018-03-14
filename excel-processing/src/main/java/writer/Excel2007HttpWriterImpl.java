package writer;


/**
 * @version: v1.0
 * @author: zhulong
 * project: order
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/3/1 0001 下午 4:48
 * modifyTime:
 * modifyBy:
 */
public class Excel2007HttpWriterImpl extends AbstractExcel2007HttpWriter {


   /* */

    /**
     * @param
     * @throws Exception
     *//*
    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub
        System.out.println("............................");
        long start = System.currentTimeMillis();
        //构建excel2007写入器
        AbstractExcel2007HttpWriter excel07Writer = new Excel2007HttpWriterImpl();
        //调用处理方法
        //
        long end = System.currentTimeMillis();
        List<ExcelToolBean> columnNameList = new ArrayList<>();
        ExcelToolBean excelToolBean = new ExcelToolBean();
        excelToolBean.setPropertyChinseName("alklka");
        excelToolBean.setPropertyName("aa");
        columnNameList.add(excelToolBean);
        List<Map<String, Object>> dataList = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("aa", "亲爱的老婆：\n" +
                "相亲相爱，鸿福满堂。\n" +
                "Love You~每天都多一点点……\n" +
                "                                小醉&Sugar");
        dataList.add(map);
        excel07Writer.process(dataList, columnNameList, null, true, "test");


        System.out.println("....................." + (end - start) / 1000);
    }*/



}
