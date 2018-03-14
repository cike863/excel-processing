package writer;

import common.ExcelToolBean;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.ReflectionUtils;
import util.DateTimeUtil;

import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

public class ExcelExportUtil {

    private static Logger logger = LoggerFactory.getLogger(ExcelExportUtil.class);

    // private static SXSSFWorkbook wb = new SXSSFWorkbook(1000);

    //private static XSSFSheet sheet = null;
    //SXSSFWorkbook

   /* public ExcelExportUtil(SXSSFWorkbook wb, XSSFSheet sheet) {
        this.wb = wb;
        this.sheet = sheet;
    }*/

    /**
     * 导出一个文件里面一个excel表
     *
     * @param dataList
     * @param columnNameList
     * @param response
     * @param closeResponse
     * @param tableName
     * @throws IOException
     */
    public static void toExcel(List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList,
                               HttpServletResponse response, boolean closeResponse,
                               String tableName, String firstRowTitle) throws IOException {
        if (dataList == null || dataList.isEmpty()) {
            logger.info("要导出的数据列为空");
            return;
        }
        if (columnNameList == null || columnNameList.isEmpty()) {
            logger.info("要导出的数据表头行为空");
            return;
        }
        if (StringUtils.isBlank(tableName)) {
            logger.info("要导出的表名为空");
            return;
        }
        logger.info("getexcel start");
        //XSSFWorkbook workBook = new XSSFWorkbook();
        SXSSFWorkbook workBook = new SXSSFWorkbook(5000);

        OutputStream outputStream = null;
        // PrintWriter out = response.getWriter();
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

            // 构建表头
            xwriteData(workBook, dataList, columnNameList, tableName, firstRowTitle);
            // 写进文档
            workBook.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            logger.error("jftj/genexcel Exception", e);
        } finally {
            if (outputStream != null) {
                if (closeResponse) {
                    outputStream.flush();
                } else {
                    outputStream.close();
                }
            }
            logger.info("getexcel end");
        }

    }

    /**
     * @Description: 通过BeanList生成excel
     * @author lizhang
     * @date 2017年5月5日 下午5:44:09
     */
    public static <T> void toExcelByListBean(List<T> dataList, List<ExcelToolBean> columnNameList,
                                             HttpServletResponse response, boolean closeResponse,
                                             String tableName, String firstRowTitle) throws IOException {
        if (dataList == null || dataList.isEmpty()) {
            logger.info("要导出的数据列为空");
            return;
        }
        if (columnNameList == null || columnNameList.isEmpty()) {
            logger.info("要导出的数据表头行为空");
            return;
        }
        if (StringUtils.isBlank(tableName)) {
            logger.info("要导出的表名为空");
            return;
        }
        logger.info("getexcel start");
        //XSSFWorkbook workBook = new XSSFWorkbook();
        SXSSFWorkbook workBook = new SXSSFWorkbook(5000);

        OutputStream outputStream = null;
        // PrintWriter out = response.getWriter();
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

            // 构建表头
            xwriteBeanData(workBook, dataList, columnNameList, tableName, firstRowTitle);
            // 写进文档
            workBook.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            logger.error("jftj/genexcel Exception", e);
        } finally {
            if (outputStream != null) {
                if (closeResponse) {
                    outputStream.flush();
                } else {
                    outputStream.close();
                }
            }
            logger.info("getexcel end");
        }

    }

    /**
     * 导出一个文件里面一个excel表
     *
     * @param dataList
     * @param columnNameList
     * @param filePath
     * @param closeResponse
     * @param tableName
     * @throws IOException
     */
    public static FileExcelResultVo toExcelFile(List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList,
                                                String filePath, boolean closeResponse,
                                                String tableName, String firstRowTitle) throws IOException {
        if (dataList == null || dataList.isEmpty()) {
            logger.info("要导出的数据列为空");
            return null;
        }
        if (columnNameList == null || columnNameList.isEmpty()) {
            logger.info("要导出的数据表头行为空");
            return null;
        }
        if (StringUtils.isBlank(tableName)) {
            logger.info("要导出的表名为空");
            return null;
        }
        logger.info("toExcelFile start");
        //XSSFWorkbook workBook = new XSSFWorkbook();
        SXSSFWorkbook workBook = new SXSSFWorkbook(5000);
        FileExcelResultVo fileExcelResultVo = new FileExcelResultVo();
        // PrintWriter out = response.getWriter();
        String fileName = new String((tableName + DateTimeUtil.buildLocalDateTimeToString_YYYYMMDDHHMMSS(LocalDateTime.now()) + (int) (Math.random() * 100)));
        filePath = filePath.concat(fileName + ".xlsx");

        fileExcelResultVo.setFileName(fileName);
        fileExcelResultVo.setFilePath(filePath);

        FileOutputStream outputStream = new FileOutputStream(filePath);
        try {
            // 构建表头
            xwriteData(workBook, dataList, columnNameList, tableName, firstRowTitle);
            // 写进文档
            workBook.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            fileExcelResultVo = null;
            logger.error("jftj/genexcel Exception", e);
        } finally {
            if (outputStream != null) {
                if (closeResponse) {
                    outputStream.flush();
                } else {
                    outputStream.close();
                }
            }
            logger.info("toExcelFile end");
        }
        return fileExcelResultVo;

    }

    /**
     * 导出一个文件里面多个excel表
     *
     * @param dataLists
     * @param columnNameList
     * @param response
     * @param closeResponse
     * @param tableName
     * @throws IOException
     */
    public static void toExcelSheets(List<List<Map<String, Object>>> dataLists,
                                     List<ExcelToolBean> columnNameList, HttpServletResponse response,
                                     boolean closeResponse, String tableName, String firstRowTitle) throws IOException {
        if (dataLists == null || dataLists.isEmpty()) {
            logger.info("要导出的数据列为空");
            return;
        }
        if (columnNameList == null || columnNameList.isEmpty()) {
            logger.info("要导出的数据表头行为空");
            return;
        }
        if (StringUtils.isBlank(tableName)) {
            logger.info("要导出的表名为空");
            return;
        }
        logger.info("getexcel start");
        //XSSFWorkbook workBook = new XSSFWorkbook();
        // 在workbook中添加一个sheet,对应Excel文件中的sheet
        // PrintWriter out = response.getWriter();
        OutputStream outputStream = null;
        SXSSFWorkbook wb = new SXSSFWorkbook(1000);
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
            int i = 0;
            for (List<Map<String, Object>> dataList : dataLists) {
                if (dataList == null || dataList.isEmpty()) {
                    continue;
                }

                writeDate(wb, dataList, columnNameList, tableName + (i++), firstRowTitle);
            }
            // 写进文档
            wb.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            logger.error("jftj/genexcel Exception", e);
        } finally {
            if (outputStream != null) {
                if (closeResponse) {
                    outputStream.flush();
                } else {
                    outputStream.close();
                }
            }
            logger.info("getexcel end");
        }
    }

    /**
     * 导出一个文件里面多个excel表
     *
     * @param dataMaps
     * @param columnNameList
     * @param response
     * @param closeResponse
     * @param tableName
     * @throws IOException
     */
    public static void toExcelSheets(Map<String, List<Map<String, Object>>> dataMaps,
                                     List<ExcelToolBean> columnNameList, HttpServletResponse response,
                                     boolean closeResponse, String tableName, String firstRowTitle) throws IOException {
        if (dataMaps == null || dataMaps.isEmpty()) {
            logger.info("要导出的数据列为空");
            return;
        }
        if (columnNameList == null || columnNameList.isEmpty()) {
            logger.info("要导出的数据表头行为空");
            return;
        }
        if (StringUtils.isBlank(tableName)) {
            logger.info("要导出的表名为空");
            return;
        }
        logger.info("getexcel start");
        //SXSSFWorkbook workBook = new XSSFWorkbook();
        // 在workbook中添加一个sheet,对应Excel文件中的sheet
        // PrintWriter out = response.getWriter();
        OutputStream outputStream = null;
        SXSSFWorkbook wb = new SXSSFWorkbook(1000);
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

            int i = 0;//防止重复的表名
            for (Map.Entry<String, List<Map<String, Object>>> entry : dataMaps.entrySet()) {
                String key = entry.getKey();
                List<Map<String, Object>> dataList = entry.getValue();
                if (dataList == null || dataList.isEmpty()) {
                    continue;
                }
                writeDate(wb, dataList, columnNameList, tableName + "_" + (i++) + key, firstRowTitle);
            }
            // 写进文档
            wb.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception e) {
            logger.error("jftj/genexcel Exception", e);
        } finally {
            if (outputStream != null) {
                if (closeResponse) {
                    outputStream.flush();
                } else {
                    outputStream.close();
                }
            }
            logger.info("getexcel end");
        }
    }

    /**
     * 创建多个excel表数据
     *
     * @param workBook
     * @param dataList
     * @param columnNameList
     * @param tableName
     */
    public static void writeDate(SXSSFWorkbook workBook, List<Map<String, Object>> dataList,
                                 List<ExcelToolBean> columnNameList, String tableName, String firstRowTitle) {
        // 在workbook中添加一个sheet,对应Excel文件中的sheet
        SXSSFSheet sheet = workBook.createSheet(tableName);
        CellStyle bigHeadStyle = workBook.createCellStyle();
        bigHeadStyle.cloneStyleFrom(XSSFCellStyleUtil.getBigHeadStyle(workBook));
        CellStyle headStyle = workBook.createCellStyle();
        headStyle.cloneStyleFrom(XSSFCellStyleUtil.getHeadStyle(workBook));
        CellStyle bodyStyle = workBook.createCellStyle();
        bodyStyle.cloneStyleFrom(XSSFCellStyleUtil.getBodyStyle(workBook));

        int size = columnNameList.size();

        // 构建表头
        SXSSFRow headRow = sheet.createRow(0);
        SXSSFCell cell = null;
        cell = headRow.createCell(0);
        headRow.setHeightInPoints(40);
        // 第一行标题大表头
        cell.setCellValue(firstRowTitle);
        cell.setCellStyle(bigHeadStyle);
        sheet.setColumnWidth(0, 10 * 256);
        // 设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
        // CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
        // 这里添加合并单元格
        CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, size);
        sheet.addMergedRegion(cellRangeAddress);

        // 第二行小表头
        SXSSFRow secondRow = sheet.createRow(1);
        cell = secondRow.createCell(0);
        cell.setCellStyle(headStyle);
        cell.setCellValue("序号");
        secondRow.setHeight((short) 500);

        for (int i = 0; i < size; i++) {
            cell = secondRow.createCell(i + 1);
            // headStyle.set
            sheet.setColumnWidth(i + 1, columnNameList.get(i).getSize() * 256);
            cell.setCellStyle(headStyle);
            cell.setCellValue(columnNameList.get(i).getPropertyChinseName());
        }
        // 构建表体数据
        for (int m = 0, dataListSize = dataList.size(); m < dataListSize; m++) {
            // i=0;
            SXSSFRow bodyRow = sheet.createRow(m + 2);
            cell = bodyRow.createCell(0);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(m + 1);
            for (int n = 0, columnSize = columnNameList.size(); n < columnSize; n++) {
                Object data = dataList.get(m).get(columnNameList.get(n).getPropertyName());
                cell = bodyRow.createCell(n + 1);
                cell.setCellStyle(bodyStyle);
               /* cell.setCellValue(data == null ? "" : data.toString());*/
                /*cell.setCellValue(data == null ? "" : data);*/
                if (data != null) {
                    if (data instanceof Double) {
                        cell.setCellValue((double) data);
                    } else if (data instanceof Integer) {
                        cell.setCellValue((int) data);
                    } else {
                        cell.setCellValue(data.toString());
                    }
                } else {
                    cell.setCellValue("");
                }
            }
        }
    }

    /**
     * 创建多个excel表数据 使用缓存区导数据
     *
     * @param workBook
     * @param dataList
     * @param columnNameList
     * @param tableName
     * @param firstRowTitle
     */
    public static void xwriteData(SXSSFWorkbook workBook, List<Map<String, Object>> dataList,
                                  List<ExcelToolBean> columnNameList, String tableName, String firstRowTitle) {

        SXSSFSheet sheet = workBook.createSheet(tableName);
        CellStyle bigHeadStyle = workBook.createCellStyle();
        bigHeadStyle.cloneStyleFrom(XSSFCellStyleUtil.getBigHeadStyle(workBook));
        CellStyle headStyle = workBook.createCellStyle();
        headStyle.cloneStyleFrom(XSSFCellStyleUtil.getHeadStyle(workBook));
        CellStyle bodyStyle = workBook.createCellStyle();
        bodyStyle.cloneStyleFrom(XSSFCellStyleUtil.getBodyStyle(workBook));

        int size = columnNameList.size();

        // 构建表头
        SXSSFRow headRow = sheet.createRow(0);
        SXSSFCell cell = null;
        cell = headRow.createCell(0);
        headRow.setHeightInPoints(40);
        // 第一行标题大表头
        cell.setCellValue(firstRowTitle);
        cell.setCellStyle(bigHeadStyle);
        sheet.setColumnWidth(0, 10 * 256);

        // 设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
        // CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
        // 这里添加合并单元格
        CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, size);
        sheet.addMergedRegion(cellRangeAddress);

        // 第二行小表头
        SXSSFRow secondRow = sheet.createRow(1);
        cell = secondRow.createCell(0);
        cell.setCellStyle(headStyle);
        cell.setCellValue("序号");
        secondRow.setHeight((short) 500);

        for (int i = 0; i < size; i++) {
            cell = secondRow.createCell(i + 1);
            // headStyle.set
            sheet.setColumnWidth(i + 1, columnNameList.get(i).getSize() * 256);
            cell.setCellStyle(headStyle);
            cell.setCellValue(columnNameList.get(i).getPropertyChinseName());
        }
        // 构建表体数据
        for (int m = 0, dataListSize = dataList.size(); m < dataListSize; m++) {
            // i=0;
            SXSSFRow bodyRow = sheet.createRow(m + 2);
            cell = bodyRow.createCell(0);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(m + 1);
            for (int n = 0, columnSize = columnNameList.size(); n < columnSize; n++) {
                Object data = dataList.get(m).get(columnNameList.get(n).getPropertyName());
                cell = bodyRow.createCell(n + 1);
                cell.setCellStyle(bodyStyle);
               /* cell.setCellValue(data == null ? "" : data.toString());*/
                if (data != null) {
                    if (data instanceof Double) {
                        cell.setCellValue((double) data);
                    } else if (data instanceof Integer) {
                        cell.setCellValue((int) data);
                    } else {
                        cell.setCellValue(data.toString());
                    }
                } else {
                    cell.setCellValue("");
                }
            }
        }
    }

    /**
     * @Description: 根据bean写excel
     * @author lizhang
     * @date 2017年5月5日 下午5:28:45
     */
    public static <T> void xwriteBeanData(SXSSFWorkbook workBook, List<T> dataList,
                                          List<ExcelToolBean> columnNameList, String tableName, String firstRowTitle) {

        // 在workbook中添加一个sheet,对应Excel文件中的sheet
        SXSSFSheet sheet = workBook.createSheet(tableName);
        CellStyle bigHeadStyle = workBook.createCellStyle();
        bigHeadStyle.cloneStyleFrom(XSSFCellStyleUtil.getBigHeadStyle(workBook));
        CellStyle headStyle = workBook.createCellStyle();
        headStyle.cloneStyleFrom(XSSFCellStyleUtil.getHeadStyle(workBook));
        CellStyle bodyStyle = workBook.createCellStyle();
        bodyStyle.cloneStyleFrom(XSSFCellStyleUtil.getBodyStyle(workBook));

        int size = columnNameList.size();

        // 构建表头
        SXSSFRow headRow = sheet.createRow(0);
        SXSSFCell cell = null;
        cell = headRow.createCell(0);
        headRow.setHeightInPoints(40);
        // 第一行标题大表头
        cell.setCellValue(firstRowTitle);
        cell.setCellStyle(bigHeadStyle);
        sheet.setColumnWidth(0, 10 * 256);

        // 设置一个合并单元格区域，使用上下左右定义CellRangeAddress区域
        // CellRangeAddress四个参数为：起始行，结束行，起始列，结束列
        // 这里添加合并单元格
        CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, size);
        sheet.addMergedRegion(cellRangeAddress);

        // 第二行小表头
        SXSSFRow secondRow = sheet.createRow(1);
        cell = secondRow.createCell(0);
        cell.setCellStyle(headStyle);
        cell.setCellValue("序号");
        secondRow.setHeight((short) 500);

        for (int i = 0; i < size; i++) {
            cell = secondRow.createCell(i + 1);
            // headStyle.set
            sheet.setColumnWidth(i + 1, columnNameList.get(i).getSize() * 256);
            cell.setCellStyle(headStyle);
            cell.setCellValue(columnNameList.get(i).getPropertyChinseName());
        }
        // 构建表体数据
        for (int m = 0, dataListSize = dataList.size(); m < dataListSize; m++) {
            // i=0;
            SXSSFRow bodyRow = sheet.createRow(m + 2);
            cell = bodyRow.createCell(0);
            cell.setCellStyle(bodyStyle);
            cell.setCellValue(m + 1);
            T dataBean = dataList.get(m);
            Map<String, Method> getMethodMap = BeanMethodReflection.initAndGetReflection(dataBean.getClass());
            for (int n = 0, columnSize = columnNameList.size(); n < columnSize; n++) {
                Method method = getMethodMap.get(columnNameList.get(n).getPropertyName());
                Object data = ReflectionUtils.invokeMethod(method, dataBean);
                cell = bodyRow.createCell(n + 1);
                cell.setCellStyle(bodyStyle);
                /* cell.setCellValue(data == null ? "" : data.toString());*/
                if (data != null) {
                    if (data instanceof Double) {
                        cell.setCellValue((double) data);
                    } else if (data instanceof Integer) {
                        cell.setCellValue((int) data);
                    } else {
                        cell.setCellValue(data.toString());
                    }
                } else {
                    cell.setCellValue("");
                }
            }
        }
    }


}
