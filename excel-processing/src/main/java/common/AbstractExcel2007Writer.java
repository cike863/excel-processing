package common;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

import static common.SpreadsheetWriter.substitute;

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
    private static SpreadsheetWriter sw;

    /**
     * 写入电子表格的主要流程
     *
     * @param fileName
     * @throws Exception
     */
    public void process(String fileName) throws Exception {
        // 建立工作簿和电子表格对象
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("sheet1");
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
        generate(sw);
        fw.close();

        // 使用产生的数据替换模板
        File templateFile = new File("template.xlsx");
        FileOutputStream out = new FileOutputStream(fileName);
        substitute(templateFile, tmp, sheetRef.substring(1), out);
        out.close();
        //删除文件之前调用一下垃圾回收器，否则无法删除模板文件
        System.gc();
        // 删除临时模板文件
        if (templateFile.isFile() && templateFile.exists()) {
            templateFile.delete();
        }
    }

    /**
     * 类使用者应该使用此方法进行写操作
     *
     * @throws Exception
     */
    public abstract void generate() throws Exception;
    /**
     * 类使用者应该使用此方法进行写操作
     *
     * @throws Exception
     * @param sw
     */
    public abstract void generate(SpreadsheetWriter sw) throws Exception;

    public static void beginSheet() throws IOException {
        sw.beginSheet();
    }

    public static void insertRow(int rowNum) throws IOException {
        sw.insertRow(rowNum);
    }

    public static void createCell(int columnIndex, String value) throws IOException {
        sw.createCell(columnIndex, value, -1);
    }

    public static void createCell(int columnIndex, double value) throws IOException {
        sw.createCell(columnIndex, value, -1);
    }

    public static void endRow() throws IOException {
        sw.endRow();
    }

    public static void endSheet() throws IOException {
        sw.endSheet();
    }


    /**
     * 在写入器中写入电子表格
     *//*
    public static class SpreadsheetWriter {

    }*/
}