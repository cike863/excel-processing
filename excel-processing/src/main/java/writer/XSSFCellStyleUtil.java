package writer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;

/**
 * @version: v1.0
 * @author: Administrator
 * @Description: project: sysmanager
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/6/29 0029 下午 12:30
 * modifyTime:
 * modifyBy:
 */
public class XSSFCellStyleUtil {
    /**
     * 合并单元格后给合并后的单元格加边框
     *
     * @param region
     * @param cs
     */
    public void setRegionStyle(CellRangeAddress region, XSSFCellStyle cs, XSSFSheet sheet) {

        int toprowNum = region.getFirstRow();
        for (int i = toprowNum; i <= region.getLastRow(); i++) {
            XSSFRow row = sheet.getRow(i);
            for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                XSSFCell cell = row.getCell(j);// XSSFCellUtil.getCell(row,
                // (short) j);
                cell.setCellStyle(cs);
            }
        }
    }

    /**
     * 大表头样式定义
     *
     * @return
     */
    public static XSSFCellStyle getBigHeadStyle(SXSSFWorkbook wb) {
        // 创建单元格样式
        XSSFCellStyle cellStyle = wb.getXSSFWorkbook().createCellStyle();
        // 设置单元格的背景颜色为淡蓝色
        // 设置单元格的背景颜色为淡蓝色
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 设置单元格居中对齐
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 设置单元格垂直居中对齐
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 创建单元格内容显示不下时自动换行
        cellStyle.setWrapText(true);
        // 设置单元格字体样式
        XSSFFont font = wb.getXSSFWorkbook().createFont();
        // 设置字体加粗
        font.setBold(true);
        //font.setFontName("宋体");
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 24);

        cellStyle.setFont(font);
        // 设置单元格边框为细线条
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        return cellStyle;
    }

    /**
     * 设置小头的单元格样式
     *
     * @return
     */
    public static XSSFCellStyle getHeadStyle(SXSSFWorkbook wb) {
        // 创建单元格样式
        XSSFCellStyle cellStyle = wb.getXSSFWorkbook().createCellStyle();
        // 设置单元格的背景颜色为淡蓝色
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 设置单元格居中对齐
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 设置单元格垂直居中对齐
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 创建单元格内容显示不下时自动换行
        cellStyle.setWrapText(true);
        // 设置单元格字体样式
        XSSFFont font = wb.getXSSFWorkbook().createFont();
        // 设置字体加粗
        font.setBold(true);
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 14);
        cellStyle.setFont(font);

        // 设置单元格边框为细线条
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        return cellStyle;
    }

    /**
     * 设置表体的单元格样式
     *
     * @return
     */
    public static XSSFCellStyle getBodyStyle(SXSSFWorkbook wb) {
        // 创建单元格样式
        XSSFCellStyle cellStyle = wb.getXSSFWorkbook().createCellStyle();
        // 设置单元格居中对齐
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 设置单元格垂直居中对齐
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 创建单元格内容显示不下时自动换行
        cellStyle.setWrapText(true);
        // 设置单元格字体样式
        XSSFFont font = wb.getXSSFWorkbook().createFont();
        // 设置字体加粗
        //font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        //font.setFontName("宋体");
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 10);
        cellStyle.setFont(font);
        // 设置单元格边框为细线条
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        return cellStyle;
    }

    /**
     * 大表头样式定义
     *
     * @return
     */
    public static XSSFCellStyle getOrderDataBigHeadStyle(SXSSFWorkbook wb) {
        // 创建单元格样式
        XSSFCellStyle cellStyle = wb.getXSSFWorkbook().createCellStyle();
        // 设置单元格的背景颜色为淡蓝色
        // 设置单元格的背景颜色为淡蓝色
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 设置单元格居中对齐
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 设置单元格垂直居中对齐
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 创建单元格内容显示不下时自动换行
        cellStyle.setWrapText(true);
        // 设置单元格字体样式
        XSSFFont font = wb.getXSSFWorkbook().createFont();
        // 设置字体加粗
        font.setBold(true);
        //font.setFontName("宋体");
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 20);

        cellStyle.setFont(font);
        // 设置单元格边框为细线条
       /* cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);*/
        return cellStyle;
    }


}
