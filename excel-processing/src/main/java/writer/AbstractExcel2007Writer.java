package writer;

import common.ExcelToolBean;
import common.SpreadsheetWriter;
import common.XMLEncoder;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import util.RegexUtil;

import java.io.IOException;
import java.util.HashMap;
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
public abstract class AbstractExcel2007Writer {
    public SpreadsheetWriter sw;
    public static final Logger LOGGER = LoggerFactory.getLogger(AbstractExcel2007Writer.class);

    /**
     * 类使用者应该使用此方法进行写操作
     *
     * @param Workbook
     * @param dataList
     * @param columnNameList
     * @param sw
     * @throws Exception
     */
    public void generate(XSSFWorkbook Workbook, List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList, SpreadsheetWriter sw) throws Exception {
        Map<String, XSSFCellStyle> styles = createStyles(Workbook);
        //电子表格开始
        sw.beginSheet();
        //第一行开始
        sw.insertRow(0);
        sw.createCell(0, "序号");
        for (int i = 0, size = columnNameList.size(); i < size; i++) {
            sw.createCell(i + 1, columnNameList.get(i).getPropertyChinseName(), 8);
        }
        sw.endRow();
        //第一行结束
        writeDataValue(1, dataList, columnNameList, sw, styles);
        //电子表格结束
        sw.endSheet();
    }


    /**
     * 创建Excel样式
     *
     * @param wb
     * @return
     */
    public static Map<String, XSSFCellStyle> createStyles(XSSFWorkbook wb) {
        Map<String, XSSFCellStyle> stylesMap = new HashMap<>();
        XSSFDataFormat fmt = wb.createDataFormat();

        XSSFCellStyle cellStringStyle = wb.createCellStyle();
        cellStringStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStringStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStringStyle.setWrapText(true);
        cellStringStyle.setBorderLeft(BorderStyle.THIN);
        cellStringStyle.setBorderBottom(BorderStyle.THIN);
        cellStringStyle.setBorderRight(BorderStyle.THIN);
        cellStringStyle.setBorderTop(BorderStyle.THIN);
        XSSFFont font = wb.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 11);
        cellStringStyle.setFont(font);
        stylesMap.put(XSSFCellStyleMap.CELL_STRING, cellStringStyle);


        XSSFCellStyle cellLongStyle = wb.createCellStyle();
        cellLongStyle.setDataFormat(fmt.getFormat("0"));
        cellLongStyle.setAlignment(HorizontalAlignment.CENTER);
        cellLongStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellLongStyle.setWrapText(true);
        cellLongStyle.setBorderLeft(BorderStyle.THIN);
        cellLongStyle.setBorderBottom(BorderStyle.THIN);
        cellLongStyle.setBorderRight(BorderStyle.THIN);
        cellLongStyle.setBorderTop(BorderStyle.THIN);
        cellLongStyle.setFont(font);
        stylesMap.put(XSSFCellStyleMap.CELL_LONG, cellLongStyle);

        XSSFCellStyle cellDoubleStyle = wb.createCellStyle();
        cellDoubleStyle.setDataFormat(fmt.getFormat("0.00"));
        cellDoubleStyle.setAlignment(HorizontalAlignment.RIGHT);
        cellDoubleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellDoubleStyle.setWrapText(true);
        cellDoubleStyle.setBorderLeft(BorderStyle.THIN);
        cellDoubleStyle.setBorderBottom(BorderStyle.THIN);
        cellDoubleStyle.setBorderRight(BorderStyle.THIN);
        cellDoubleStyle.setBorderTop(BorderStyle.THIN);
        cellDoubleStyle.setFont(font);
        stylesMap.put(XSSFCellStyleMap.CELL_DOUBLE, cellDoubleStyle);

        XSSFCellStyle cellDateStyle = wb.createCellStyle();
        cellDateStyle.setDataFormat(fmt.getFormat("yyyy-MM-dd"));
        cellDateStyle.setAlignment(HorizontalAlignment.RIGHT);
        cellDateStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellDateStyle.setWrapText(true);
        cellDateStyle.setBorderLeft(BorderStyle.THIN);
        cellDateStyle.setBorderBottom(BorderStyle.THIN);
        cellDateStyle.setBorderRight(BorderStyle.THIN);
        cellDateStyle.setBorderTop(BorderStyle.THIN);
        cellDateStyle.setFont(font);
        stylesMap.put(XSSFCellStyleMap.CELL_DATE, cellDateStyle);

        XSSFCellStyle colTitleStyle = wb.createCellStyle();
        // 设置单元格的背景颜色为淡蓝色
        colTitleStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        colTitleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        colTitleStyle.setAlignment(HorizontalAlignment.CENTER);
        colTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        colTitleStyle.setWrapText(true);
        colTitleStyle.setBorderLeft(BorderStyle.THIN);
        colTitleStyle.setBorderBottom(BorderStyle.THIN);
        colTitleStyle.setBorderRight(BorderStyle.THIN);
        colTitleStyle.setBorderTop(BorderStyle.THIN);
        colTitleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //字体
        XSSFFont font5 = wb.createFont();
        font5.setFontName("微软雅黑");
        font5.setFontHeightInPoints((short) 13);
        colTitleStyle.setFont(font5);
        stylesMap.put(XSSFCellStyleMap.COL_TITLE, colTitleStyle);

        XSSFCellStyle sheetTitleStyle = wb.createCellStyle();
        // 设置单元格的背景颜色为淡蓝色
        sheetTitleStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
        sheetTitleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        sheetTitleStyle.setAlignment(HorizontalAlignment.CENTER);
        sheetTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        sheetTitleStyle.setWrapText(true);
        sheetTitleStyle.setBorderLeft(BorderStyle.THIN);
        sheetTitleStyle.setBorderBottom(BorderStyle.THIN);
        sheetTitleStyle.setBorderRight(BorderStyle.THIN);
        sheetTitleStyle.setBorderTop(BorderStyle.THIN);
        //字体
        XSSFFont font6 = wb.createFont();
        font6.setFontName("微软雅黑");
        font6.setFontHeightInPoints((short) 24);
        // 设置字体加粗
        //font6.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font6.setBold(Boolean.TRUE);
        sheetTitleStyle.setFont(font6);
        stylesMap.put(XSSFCellStyleMap.SHEET_TITLE, sheetTitleStyle);
        return stylesMap;
    }


    /**
     * 写数据区域
     *
     * @param indexRowNum    行号
     * @param dataList       数据
     * @param columnNameList title list
     * @param sw             sw
     * @param stylesMap      样式
     * @return
     * @throws IOException
     */
    public int writeDataValue(int indexRowNum, List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList, SpreadsheetWriter sw, Map<String, XSSFCellStyle> stylesMap) throws IOException {
        int rownumSize = dataList.size();
        for (int rownum = 0; rownum < rownumSize; rownum++) {
            //插入新行
            sw.insertStyleRow(indexRowNum, 20);
            //建立新单元格,索引值从0开始,表示第一列
            sw.createCell(0, indexRowNum - 1, stylesMap.get(XSSFCellStyleMap.CELL_LONG).getIndex());
            for (int n = 0, columnSize = columnNameList.size(); n < columnSize; n++) {
                Object data = dataList.get(rownum).get(columnNameList.get(n).getPropertyName());
                if (data != null) {
                    if (data instanceof Double) {
                        sw.createCell(n + 1, (double) data, stylesMap.get(XSSFCellStyleMap.CELL_DOUBLE).getIndex());
                    } else if (data instanceof Integer) {
                        sw.createCell(n + 1, (int) data, stylesMap.get(XSSFCellStyleMap.CELL_LONG).getIndex());
                    } else {
                        if (StringUtils.isNotBlank(data.toString())) {
                            if (RegexUtil.validText(data.toString())) {
                                sw.createCell(n + 1, data.toString(), stylesMap.get(XSSFCellStyleMap.CELL_STRING).getIndex());
                            } else {
                                sw.createCell(n + 1, "<![CDATA[" + data.toString() + "]]>", stylesMap.get(XSSFCellStyleMap.CELL_STRING).getIndex());
                            }
                        } else {
                            sw.createCell(n + 1, " ", stylesMap.get(XSSFCellStyleMap.CELL_STRING).getIndex());
                        }
                    }
                } else {
                    sw.createCell(n + 1, " ", stylesMap.get(XSSFCellStyleMap.CELL_STRING).getIndex());
                }

            }
            //结束行
            sw.endRow();
            indexRowNum++;
            if (indexRowNum % 100 == 0) {
                sw.flush();
            }
        }
        return indexRowNum;
    }

    /**
     * 写表头
     *
     * @param columnNameList title list
     * @param sw             sw
     * @param tableName      表明
     * @param styles         样式
     * @throws Exception
     */
    public void writeTitle(List<ExcelToolBean> columnNameList, SpreadsheetWriter sw, String tableName, Map<String, XSSFCellStyle> styles, String workbookEncoding) throws Exception {
        XSSFCellStyle colStyle = styles.get(XSSFCellStyleMap.COL_TITLE);
        XSSFCellStyle titleStyle = styles.get(XSSFCellStyleMap.SHEET_TITLE);

        //电子表格开始
        if (StringUtils.isBlank(workbookEncoding)) {
            sw.beginSheetStyle();
        } else {
            sw.beginSheetStyle(workbookEncoding);
        }
        sw.defaultRowHeight(null);
        sw.beginColsStyle();
        int size = columnNameList.size();
        sw.defaultColsStyle(1, 10);
        for (int i = 0; i < size; i++) {
            sw.defaultColsStyle(i + 2, columnNameList.get(i).getSize());
        }
        sw.endColsStyle();
        sw.beginSheetData();
        //表头
        sw.insertStyleRow(0, 40);
        sw.createCell(0, tableName, titleStyle.getIndex());
        for (int i = 1; i < size; i++) {
            sw.createCellNotValue(i + 1, -1);
        }
        //结束行
        sw.endRow();

        //第一行开始
        sw.insertStyleRow(1, 30);
        sw.createCell(0, "序号", colStyle.getIndex());
        for (int i = 0; i < size; i++) {
            sw.createCell(i + 1, columnNameList.get(i).getPropertyChinseName(), colStyle.getIndex());
        }
        sw.endRow();
        //第一行结束
    }

    /**
     * 结束表
     *
     * @param size 合并单元格的大小
     * @param sw   sw
     * @throws Exception
     */
    public void sheelEnd(int size, SpreadsheetWriter sw) throws Exception {
        sw.endSheetData();
        sw.beginMergeCells(1);
        String ref = new CellReference(0, size).formatAsString().replaceAll("\\$", "");
        sw.insertMergeCell("A1:" + ref);
        sw.endMergeCells();
        sw.endWorksheet();
    }

    /**
     * 写数据
     *
     * @param wb             XSSFWorkbook
     * @param dataList       数据
     * @param columnNameList 抬头
     * @param sw             sw
     * @param tableName      表明
     * @throws Exception
     */
    public void generate(XSSFWorkbook wb, List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList, SpreadsheetWriter sw, String tableName, Map<String, XSSFCellStyle> styles) throws Exception {
        //styles= createStyles2(wb);
        XSSFCellStyle colStyle = styles.get(XSSFCellStyleMap.COL_TITLE);
        XSSFCellStyle titleStyle = styles.get(XSSFCellStyleMap.SHEET_TITLE);
        XSSFCellStyle longStyle = styles.get(XSSFCellStyleMap.CELL_LONG);
        XSSFCellStyle doubleStyle = styles.get(XSSFCellStyleMap.CELL_DOUBLE);
        XSSFCellStyle stringStyle = styles.get(XSSFCellStyleMap.CELL_STRING);

        //电子表格开始
        sw.beginSheetStyle();
        sw.defaultRowHeight(null);
        sw.beginColsStyle();
        int size = columnNameList.size();
        sw.defaultColsStyle(1, 10);
        for (int i = 0; i < size; i++) {
            sw.defaultColsStyle(i + 2, columnNameList.get(i).getSize());
        }
        sw.endColsStyle();
        sw.beginSheetData();
        //表头
        sw.insertStyleRow(0, 40);
        sw.createCell(0, tableName, titleStyle.getIndex());
        for (int i = 1; i < size; i++) {
            sw.createCellNotValue(i + 1, -1);
        }
        //结束行
        sw.endRow();

        //第一行开始
        sw.insertStyleRow(1, 30);
        sw.createCell(0, "序号", colStyle.getIndex());
        for (int i = 0; i < size; i++) {
            sw.createCell(i + 1, columnNameList.get(i).getPropertyChinseName(), colStyle.getIndex());
        }
        sw.endRow();
        //第一行结束
        String value = " ";
        int rownumSize = dataList.size();
        for (int rownum = 0; rownum < rownumSize; rownum++) {
            //插入新行
            sw.insertStyleRow(rownum + 2, 20);
            //建立新单元格,索引值从0开始,表示第一列
            sw.createCell(0, rownum + 1, longStyle.getIndex());
            for (int n = 0, columnSize = columnNameList.size(); n < columnSize; n++) {
                //Object data = dataList.get(rownum).get(columnNameList.get(n).getPropertyName());
                Object data = dataList.get(rownum).get(columnNameList.get(n).getPropertyName());
                if (data != null) {
                    if (data instanceof Double) {
                        sw.createCell(n + 1, (double) data, doubleStyle.getIndex());
                    } else if (data instanceof Integer) {
                        sw.createCell(n + 1, (int) data, longStyle.getIndex());
                    } else {
                        sw.createCell(n + 1, "<![CDATA[" + data.toString() + "]]>", stringStyle.getIndex());
                    }
                } else {
                    sw.createCell(n + 1, value);
                }

            }
            //结束行
            sw.endRow();
        }
        //电子表格结束
        //sw.endSheet();
        sw.endSheetData();
        sw.beginMergeCells(1);
        String ref = new CellReference(0, size).formatAsString().replaceAll("\\$", "");
        sw.insertMergeCell("A1:" + ref);
        sw.endMergeCells();
        sw.endWorksheet();
    }

    public void generate(List<Map<String, Object>> dataList, List<ExcelToolBean> columnNameList, SpreadsheetWriter sw, String tableName, Map<String, XSSFCellStyle> styles) throws Exception {
        XSSFCellStyle stringStyle = styles.get("cell_string");
        XSSFCellStyle longStyle = styles.get("cell_long");
        XSSFCellStyle doubleStyle = styles.get("cell_double");
        XSSFCellStyle colStyle = styles.get("col_title");
        XSSFCellStyle titleStyle = styles.get("sheet_title");

        //电子表格开始
        sw.beginSheetStyle();
        sw.defaultRowHeight(null);
        sw.beginColsStyle();
        int size = columnNameList.size();
        sw.defaultColsStyle(1, 10);
        for (int i = 0; i < size; i++) {
            sw.defaultColsStyle(i + 2, columnNameList.get(i).getSize());
        }
        sw.endColsStyle();
        sw.beginSheetData();
        //表头
        sw.insertStyleRow(0, 40);
        sw.createCell(0, tableName, titleStyle.getIndex());
        for (int i = 1; i < size; i++) {
            sw.createCellNotValue(i + 1, -1);
        }
        //结束行
        sw.endRow();

        //第一行开始
        sw.insertStyleRow(1, 30);
        sw.createCell(0, "序号", colStyle.getIndex());
        for (int i = 0; i < size; i++) {
            sw.createCell(i + 1, columnNameList.get(i).getPropertyChinseName(), colStyle.getIndex());
        }
        sw.endRow();
        //第一行结束
        String value = " ";
        int rownumSize = dataList.size();
        for (int rownum = 0; rownum < rownumSize; rownum++) {
            //插入新行
            sw.insertStyleRow(rownum + 2, 20);
            //建立新单元格,索引值从0开始,表示第一列
            sw.createCell(0, rownum + 1, longStyle.getIndex());
            for (int n = 0, columnSize = columnNameList.size(); n < columnSize; n++) {
                //Object data = dataList.get(rownum).get(columnNameList.get(n).getPropertyName());
                Object data = dataList.get(rownum).get(columnNameList.get(n).getPropertyName());
                if (data != null) {
                    if (data instanceof Double) {
                        sw.createCell(n + 1, (double) data, doubleStyle.getIndex());
                    } else if (data instanceof Integer) {
                        sw.createCell(n + 1, (int) data, longStyle.getIndex());
                    } else {
                        if (StringUtils.isNotBlank(data.toString())) {
                            boolean check = RegexUtil.text(data.toString());
                            //sw.createCell(n + 1, "<![CDATA[" + data.toString().replaceAll("\b", "").replaceAll("\u0007", "").replaceAll("\u0002", "").replaceAll(" \b", "") + "]]>", stringStyle.getIndex());
                            sw.createCell(n + 1, "<![CDATA[" + XMLEncoder.StringFilter(data.toString()) + "]]>", stringStyle.getIndex());
                        } else {
                            sw.createCell(n + 1, " ", stringStyle.getIndex());
                        }
                    }
                } else {
                    sw.createCell(n + 1, value);
                }

            }
            //结束行
            sw.endRow();
        }
        //电子表格结束
        //sw.endSheet();
        sw.endSheetData();
        sw.beginMergeCells(1);
        String ref = new CellReference(0, size).formatAsString().replaceAll("\\$", "");
        sw.insertMergeCell("A1:" + ref);
        sw.endMergeCells();
        sw.endWorksheet();
    }
}