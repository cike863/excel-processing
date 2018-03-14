package writer;

import common.ExcelToolBean;
import common.SpreadsheetWriter;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import util.DateTimeUtil;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;



/**
 * @version: v1.0
 * @author: zhulong
 * project: order
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/3/1 0001 下午 4:46
 * modifyTime:
 * modifyBy:
 */
public class AbstractExcel2007HttpWriter extends AbstractExcel2007Writer {
    /**
     * 写入电子表格的主要流程
     *
     * @param
     * @throws Exception
     */
    public void process(List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList,
                        HttpServletResponse response, boolean closeResponse, String tableName) throws Exception {
        LOGGER.info("AbstractExcel2007Writer start");
        // 建立工作簿和电子表格对象
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(tableName);
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
        //generate(wb, dataList, columnNameList, sw, tableName);
        Map<String, XSSFCellStyle> styles = createStyles(wb);
        //List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList, SpreadsheetWriter sw, String tableName, Map<String, XSSFCellStyle> styles
        generate(dataList, columnNameList, sw, tableName, styles);
        fw.close();
        OutputStream outputStream = null;
        try {
            outputStream = response.getOutputStream();
            // response.setContentType("application/dowload");
            response.reset();
            response.setContentType("application/msexcel");
            response.setHeader(
                    "Content-disposition",
                    "attachment;filename=\""
                            + new String((java.net.URLEncoder.encode(tableName + DateTimeUtil.buildLocalDateTimeToString_YYYYMMDDHHMMSS(LocalDateTime.now()) + (int) (Math.random() * 100) + ".xlsx", "UTF-8"))
                            .getBytes("UTF-8"), "GB2312") + "\"");

            // 使用产生的数据替换模板
            File templateFile = new File("template.xlsx");
            SpreadsheetWriter.substitute(templateFile, tmp, sheetRef.substring(1), outputStream);
            //删除文件之前调用一下垃圾回收器，否则无法删除模板文件
            System.gc();
            // 删除临时模板文件
           /* if (templateFile.isFile() && templateFile.exists()) {
                templateFile.delete();
            }*/
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            LOGGER.error("AbstractExcel2007Writer Exception", e);
        } finally {
            if (outputStream != null) {
                if (closeResponse) {
                    outputStream.flush();
                } else {
                    outputStream.close();
                }
            }
            LOGGER.info("AbstractExcel2007Writer end");
        }

    }

    /**
     * 写入电子表格的主要流程
     *
     * @param
     * @throws Exception
     */
    public void processTitle(List<ExcelToolBean> columnNameList, String tableName) throws Exception {
        LOGGER.info("AbstractExcel2007Writer start");
        // 建立工作簿和电子表格对象
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(tableName);
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
        //List<ExcelToolBean> columnNameList, SpreadsheetWriter sw, String tableName, Map<String, XSSFCellStyle> styles
        writeTitle(columnNameList, sw, tableName, createStyles(wb),null);
        fw.close();
    }


}