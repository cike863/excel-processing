package writer;

import common.ExcelToolBean;
import common.SpreadsheetWriter;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import util.DateTimeUtil;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.Writer;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * 输出excel到文件
 *
 * @version: v1.0
 * @author: zhulong
 * project: order
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/3/1 0001 下午 4:46
 * modifyTime:
 * modifyBy:
 */
public class AbstractExcel2007FileWriter extends AbstractExcel2007Writer {

    /**
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub
        System.out.println("............................");
        long start = System.currentTimeMillis();
        //构建excel2007写入器
        AbstractExcel2007FileWriter excel07Writer = new Excel2007FileWriterImpl();
        //调用处理方法
        //
        long end = System.currentTimeMillis();
        List<ExcelToolBean> columnNameList = new ArrayList<>();
        ExcelToolBean excelToolBean = new ExcelToolBean();
        excelToolBean.setPropertyChinseName("alklka");
        excelToolBean.setPropertyName("aa");
        excelToolBean.setSize(20);
        columnNameList.add(excelToolBean);
        List<Map<String, Object>> dataList = new ArrayList<>();
        Map<String, Object> map = new HashMap<>();
        map.put("aa", "亲爱的老婆：\n" +
                "相亲相爱，鸿福满堂。\n" +
                "Love You~每天都多一点点……\n" +
                "                                小醉&Sugar");
        dataList.add(map);
        //List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList,String filePath, String tableName
        String path = "C:/Users/Administrator/Desktop/test11/";
        excel07Writer.process(dataList, columnNameList, path, "test11");


        System.out.println("....................." + (end - start) / 1000);
    }


    /**
     * 写入电子表格的主要流程
     *
     * @param
     * @throws Exception
     */
    public FileExcelResultVo process(List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList,
                                     String filePath, String tableName) throws Exception {
        LOGGER.info("AbstractExcel2007FileWriter start");
        FileExcelResultVo fileExcelResultVo = new FileExcelResultVo();
        // 建立工作簿和电子表格对象
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(tableName);
        Map<String, XSSFCellStyle> styles = createStyles(wb);
        // 持有电子表格数据的xml文件名 例如 /xl/worksheets/sheet1.xml
        String sheetRef = sheet.getPackagePart().getPartName().getName();

        // 保存模板
        FileOutputStream os = new FileOutputStream("template.xlsx");
        wb.write(os);
        os.close();

        // 生成xml文件
        File tmp = File.createTempFile("sheet", ".xml");
        Writer fw = new FileWriter(tmp);
        sw = new SpreadsheetWriter(fw);
        //generate(wb, dataList, columnNameList,sw,tableName);
        //List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList, SpreadsheetWriter sw, String tableName, Map<String, XSSFCellStyle> styles

        generate(wb, dataList, columnNameList, sw, tableName, styles);
        fw.close();

        String fileName = new String((tableName + DateTimeUtil.buildLocalDateTimeToString_YYYYMMDDHHMMSS(LocalDateTime.now()) + (int) (Math.random() * 100)));
        filePath = filePath.concat(fileName + ".xlsx");

        fileExcelResultVo.setFileName(fileName);
        fileExcelResultVo.setFilePath(filePath);

        FileOutputStream fileOutputStream = new FileOutputStream(filePath);

        try {
            // 使用产生的数据替换模板
            File templateFile = new File("template.xlsx");
            SpreadsheetWriter.substitute(templateFile, tmp, sheetRef.substring(1), fileOutputStream);
            //删除文件之前调用一下垃圾回收器，否则无法删除模板文件
            System.gc();
            // 删除临时模板文件
            if (templateFile.isFile() && templateFile.exists()) {
                templateFile.delete();
            }
            fileOutputStream.close();
        } catch (Exception e) {
            fileExcelResultVo = null;
            LOGGER.error("AbstractExcel2007FileWriter Exception", e);
        } finally {
            if (fileOutputStream != null) {
                fileOutputStream.close();
            }
            LOGGER.info("AbstractExcel2007FileWriter end");
        }
        return fileExcelResultVo;
    }

}