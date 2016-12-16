package com.lingminjun.tools.excel;

import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;

/**
 * Created by lingminjun on 16/12/15.
 */
public class ExcelUtils {

    /**
     * 文件按照行读取,并采用split分割行数据
     * @param in
     * @param out
     * @param split
     * @throws FileNotFoundException
     */
    public static void lineTxtToExcel(String in, String out, String split) throws IOException {
        //文件按照行读取
        FileReader reader = new FileReader(in);
        BufferedReader br = new BufferedReader(reader);

        //分割符号
        String spl = split;
        if (spl == null || spl.length() == 0) {
            spl = " ";
        }

        SXSSFWorkbook wkb = null;//new XSSFWorkbook();
        SXSSFSheet sheet= null;//wkb.createSheet("未命名");

        String str = null;
        int row = 0;
        int max_row = 400000;//40万
        int file_num = 0;
        while((str = br.readLine()) != null) {
            String[] ss = str.split(spl);
            if (ss != null) {

                if (wkb == null) {
                    wkb = new SXSSFWorkbook();
                    sheet = wkb.createSheet("未命名");
                }

                //在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
                SXSSFRow ssrow = sheet.createRow(row++);
                insertRow(ssrow, ss);

                if (row >= max_row) {
                    //输出Excel文件

                    String pathStr = out;
                    if (file_num > 0) {
                        String path = out.substring(0,out.lastIndexOf('.'));
                        String ext = out.substring(out.lastIndexOf('.'));
                        pathStr = path + file_num + ext;
                    }
                    file_num++;

                    saveToFile(wkb,pathStr);
                    System.out.println("size:"+row+";save:"+pathStr);

                    row = 0;
                    wkb = null;
                    sheet = null;
                }
            }
        }

        br.close();
        reader.close();

        //输出Excel文件
        if (wkb != null) {
            String pathStr = out;
            if (file_num > 0) {
                String path = out.substring(0,out.lastIndexOf('.'));
                String ext = out.substring(out.lastIndexOf('.'));
                pathStr = path + file_num + ext;
            }
            saveToFile(wkb,pathStr);
            System.out.println("size:"+row+";save:"+pathStr);
        }
    }

    private static void saveToFile(SXSSFWorkbook wkb, String path) throws IOException {
        FileOutputStream output = new FileOutputStream(path);
        wkb.write(output);
        output.flush();
    }

    //在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
    private static void insertRow(SXSSFRow row, String[] values) {
        int size = values.length;
        for (int i = 0; i < size; i++) {
            String v = values[i];
            row.createCell(i).setCellValue(v);
        }
    }
}
