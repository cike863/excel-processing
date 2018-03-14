package common;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;

import java.io.*;
import java.util.Calendar;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * @version: v1.0
 * @author: zhulong
 * project: order
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/3/2 0002 上午 11:55
 * modifyTime:
 * modifyBy:
 */
public class SpreadsheetWriter {
    private final Writer _out;
    private int _rownum;
    private static String LINE_SEPARATOR = System.getProperty("line.separator");

    public SpreadsheetWriter(Writer out) {
        _out = out;
    }

    public void beginSheet() throws IOException {
        _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        _out.write("<sheetData>" + LINE_SEPARATOR);
    }

    public void beginSheetStyle() throws IOException {
        _out.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + LINE_SEPARATOR);
        //_out.write("<sheetData>" + LINE_SEPARATOR);
    }

    public void beginSheet(String encoding) throws IOException {
        _out.write("<?xml version=\"1.0\" encoding=\"" + encoding + "\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        _out.write("<sheetData>" + LINE_SEPARATOR);
    }

    public void beginSheetStyle(String encoding) throws IOException {
        _out.write("<?xml version=\"1.0\" encoding=\"" + encoding + "\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" + LINE_SEPARATOR);
        //_out.write("<sheetData>" + LINE_SEPARATOR);
    }

    //style="1"
    public void defaultColsStyle(int col, int width) throws IOException {
        _out.write("<col min=\"" + col + "\" max=\"" + col + "\" width=\"" + width + "\" customWidth=\"1\"/>");
    }
    /*
    public void defaultColsStyle(int col, int width) throws IOException {
        _out.write("<col min=\"" + col + "\" max=\"" + col + "\" width=\"" + width + "\"/>");
    }*/

    public void beginColsStyle() throws IOException {
        _out.write("<cols>");
    }

    public void endColsStyle() throws IOException {
        _out.write("</cols>" + LINE_SEPARATOR);
    }

    public void beginSheetData() throws IOException {
        _out.write("<sheetData>" + LINE_SEPARATOR);
    }

    public void defaultRowHeight(Integer defaultRowHeight) throws IOException {
        if (defaultRowHeight == null) {
            _out.write("<sheetFormatPr defaultRowHeight=\"20\"/>" + LINE_SEPARATOR);
        } else {
            _out.write("<sheetFormatPr defaultRowHeight=\"" + defaultRowHeight + "\"/>" + LINE_SEPARATOR);
        }

    }

    public void endSheet() throws IOException {
        _out.write("</sheetData>");
        _out.write("</worksheet>");
    }

    public void endSheetData() throws IOException {
        _out.write("</sheetData>");
    }

    public void endWorksheet() throws IOException {
        _out.write("</worksheet>");
    }

    public void endSheetStyle() throws IOException {
        _out.write("</sheetData>" + LINE_SEPARATOR);
        _out.write("</worksheet>");
    }

    /**
     * 插入新行
     *
     * @param rownum 以0开始
     */
    public void insertStyleRow(int rownum, int ht) throws IOException {
        _out.write("<row r=\"" + (rownum + 1) + "\" ht=\"" + ht + "\" customHeight=\"1\">" + LINE_SEPARATOR);
        this._rownum = rownum;
    }
    /**
     * 插入新行
     *
     * @param rownum 以0开始
     */
    /*
    public void insertStyleRow(int rownum, int ht) throws IOException {
        _out.write("<row r=\"" + (rownum + 1) + "\" ht=\"" + ht + "\">" + LINE_SEPARATOR);
        this._rownum = rownum;
    }*/

    /**
     * 开始合并单元格
     *
     * @param mergeCellsCount 以0开始
     */
    public void beginMergeCells(int mergeCellsCount) throws IOException {
        _out.write("<mergeCells count=\"" + mergeCellsCount + "\">");
    }

    /**
     * 结束合并单元格
     *
     * @param
     */
    public void endMergeCells() throws IOException {
        _out.write("</mergeCells>");
    }

    /**
     * 插入合并单元格--区域
     *
     * @param ref
     */
    public void insertMergeCell(String ref) throws IOException {
        _out.write("<mergeCell ref=\"" + ref + "\"/>");
    }

    /**
     * 插入新行
     *
     * @param rownum 以0开始
     */
    public void insertRow(int rownum) throws IOException {
        _out.write("<row r=\"" + (rownum + 1) + "\">" + LINE_SEPARATOR);
        this._rownum = rownum;
    }

    /**
     * 插入行结束标志
     */
    public void endRow() throws IOException {
        _out.write("</row>" + LINE_SEPARATOR);
    }

    /**
     * 插入新列
     *
     * @param columnIndex
     * @param value
     * @param styleIndex
     * @throws IOException
     */
    public void createCell(int columnIndex, String value, int styleIndex)
            throws IOException {
        String ref = new CellReference(_rownum, columnIndex)
                .formatAsString();
        _out.write("<c r=\"" + ref + "\" t=\"inlineStr\"");
        if (styleIndex != -1) {
            _out.write(" s=\"" + styleIndex + "\"");
        }
        _out.write(">");
        _out.write("<is><t>" + value + "</t></is>");
       /* _out.write("<is><t>" + XMLEncoder.encode(value) + "</t></is>");*/
        _out.write("</c>");
    }

    /**
     * 插入新列--没有值
     *
     * @param columnIndex
     * @param styleIndex
     * @throws IOException
     */
    public void createCellNotValue(int columnIndex, int styleIndex)
            throws IOException {
        String ref = new CellReference(_rownum, columnIndex)
                .formatAsString();
        _out.write("<c r=\"" + ref + "\" t=\"inlineStr\"");
        if (styleIndex != -1) {
            _out.write(" s=\"" + styleIndex + "\"");
        }
        _out.write(">");
        _out.write("</c>");
    }

    public void createCell(int columnIndex, String value)
            throws IOException {
        createCell(columnIndex, value, -1);
    }

    public void createCell(int columnIndex, double value, int styleIndex)
            throws IOException {
        String ref = new CellReference(_rownum, columnIndex)
                .formatAsString();
        _out.write("<c r=\"" + ref + "\" t=\"n\"");
        if (styleIndex != -1) {
            _out.write(" s=\"" + styleIndex + "\"");
        }
        _out.write(">");
        _out.write("<v>" + value + "</v>");
        _out.write("</c>");
    }

    public void createCell(int columnIndex, double value)
            throws IOException {
        createCell(columnIndex, value, -1);
    }

    public void createCell(int columnIndex, Calendar value, int styleIndex)
            throws IOException {
        createCell(columnIndex, DateUtil.getExcelDate(value, false),
                styleIndex);
    }

    /**
     * @param zipfile the template file
     * @param tmpfile the XML file with the sheet data
     * @param entry   the name of the sheet entry to substitute, e.g. xl/worksheets/sheet1.xml
     * @param out     the stream to write the result to
     */
    public static void substitute(File zipfile, File tmpfile, String entry,
                                  OutputStream out) throws IOException {
        ZipFile zip = new ZipFile(zipfile);
        ZipOutputStream zos = new ZipOutputStream(out);

        @SuppressWarnings("unchecked")
        Enumeration<ZipEntry> en = (Enumeration<ZipEntry>) zip.entries();
        while (en.hasMoreElements()) {
            ZipEntry ze = en.nextElement();
            if (!ze.getName().equals(entry)) {
                zos.putNextEntry(new ZipEntry(ze.getName()));
                InputStream is = zip.getInputStream(ze);
                copyStream(is, zos);
                is.close();
            }
        }
        zos.putNextEntry(new ZipEntry(entry));
        InputStream is = new FileInputStream(tmpfile);
        copyStream(is, zos);
        is.close();
        zos.close();
    }

    public void flush() throws IOException {
        _out.flush();
    }

    private static void copyStream(InputStream in, OutputStream out)
            throws IOException {
        byte[] chunk = new byte[1024];
        int count;
        int i = 0;
        while ((count = in.read(chunk)) >= 0) {
            if (i++ % 1000 == 0) {
                out.flush();
            }
            out.write(chunk, 0, count);
        }
    }
}
