package com.lingminjun.tools.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
        File file = new File(path);
        File parent = file.getParentFile();
        if (!parent.exists() || !parent.isDirectory()) {
            parent.mkdir();
        }
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

    private static void fillRow(SXSSFRow xrow, Row row) {
        if(row!=null&&!(row.equals(""))){
            int columns_num = row.getLastCellNum();//获取列数

            for( int columns=0;columns<columns_num;columns++){
                Cell cell = row.getCell(columns);
                if(cell!=null){
                    String value = null;
                    switch ( cell.getCellType()) {
                        case XSSFCell.CELL_TYPE_STRING: // 字符串
                            value = cell.getStringCellValue();
                            break;
                        case XSSFCell.CELL_TYPE_NUMERIC: // 数字
                            double strCell = cell.getNumericCellValue();
                            value = "" + strCell;
                            break;
                        case XSSFCell.CELL_TYPE_BLANK: // 空值
                            value = "";
                            break;
                        case XSSFCell.CELL_TYPE_BOOLEAN: // 空值
                            boolean b = cell.getBooleanCellValue();
                            value = "" + b;
                            break;
                        default:
                            value = "";
                            System.out.println("warning!! 不支持数据类型");
                            break;
                    }
                    xrow.createCell(columns).setCellValue(value);
                }
            }
        }
    }

    /**
     * 拆分excel,多sheet忽略
     * @param in
     * @param out
     * @param max
     * @throws IOException
     */
    public static void splitExcel(String in, String out, int max) throws IOException {

        String xls_read_Address=in;//读取
//        String xls_write_Address=out;//写入

        try {

//            File excel_file = new File(xls_read_Address);//读取的文件路径
//            FileInputStream input = new FileInputStream(excel_file);  //读取的文件路径
//            String fileName = "D:\\excel\\xlsx_test.xlsx";
//            XSSFWorkbook xssfWorkbook = new XSSFWorkbook( fileName);

            Workbook wb = null;
            if (in.endsWith("xlsx")) {
                wb = new XSSFWorkbook(in);
            } else {
                InputStream is = new FileInputStream(in);
                wb = new HSSFWorkbook(is);
            }

            int sheet_numbers = wb.getNumberOfSheets();//获取表的总数

            if (sheet_numbers <= 0) {
                System.out.println("没有找到 excel sheet");
                return;
            } else if (sheet_numbers > 1) {
                System.out.println("warning!! 多个sheet, 只能处理第一个");
            }

            Sheet sheet = wb.getSheetAt(0);

            SXSSFWorkbook in_wkb = null;//new XSSFWorkbook();
            SXSSFSheet in_sheet= null;//wkb.createSheet("未命名");

            int rows_num = sheet.getLastRowNum();//获取行数
            int r = 0;
            int file_num = 0;//文件个数

            while (r < rows_num) {
                Row row = sheet.getRow(r);//取得某一行   对象

                if (r%max == 0 && in_wkb != null) {
                    String pathStr = out;
                    if (file_num > 0) {
                        String path = out.substring(0,out.lastIndexOf('.'));
                        String ext = out.substring(out.lastIndexOf('.'));
                        pathStr = path + file_num + ext;
                    }
                    file_num++;

                    saveToFile(in_wkb,pathStr);
                    System.out.println("size:"+r%max+";save:"+pathStr);
                    in_wkb = null;
                    in_sheet = null;
                }

                if (in_wkb == null) {
                    in_wkb = new SXSSFWorkbook();
                    in_sheet = in_wkb.createSheet("未命名");
                }

                //开始检查是否重新创建
                fillRow(in_sheet.createRow(r%max),row);

                r++;//下一个
            }

            //输出Excel文件
            if (in_wkb != null) {
                String pathStr = out;
                if (file_num > 0) {
                    String path = out.substring(0,out.lastIndexOf('.'));
                    String ext = out.substring(out.lastIndexOf('.'));
                    pathStr = path + file_num + ext;
                }
                saveToFile(in_wkb,pathStr);
                System.out.println("size:"+r%max+";save:"+pathStr);
            }

//            input.close();
        } catch (IOException ex) {
            System.out.println("拆分报错!!!");
            ex.printStackTrace();
        }
    }
}
