package com.lingminjun.tools;

import com.lingminjun.tools.excel.ExcelUtils;

import java.io.IOException;

/**
 * Created by lingminjun on 16/12/15.
 */
public class Main {
    private static void help() {
        System.out.println(
                "--txt_to_excel IN_PUT_PATH OUT_PUT_PATH [SPLIT_STRING] 按照行读取文本并转换成excel)\n" +
                        "\tIN_PUT_PATH: 输入文本路径,不能为空\n" +
                        "\tOUT_PUT_PATH: 输出excel地址,不能为空,请自带后缀名xlsx\n" +
                        "\tSPLIT_STRING: 分割字符,不传时默认为空格分割\n" +
                        "\n");
    }

    public static void main(String[] vars) {

//        if (vars == null || vars.length == 0) {
//            help();
//            return;
//        }
//
//        String cmd = vars[0];
//        if (cmd.contains("--txt_to_excel")) {
//            try {
//                ExcelUtils.lineTxtToExcel(vars[1].trim(),vars[2].trim(),(vars.length >= 4 ? vars[3] : null));
//            } catch (Throwable e) {
//                e.printStackTrace();
//                help();
//            }
//            return;
//        }

        String in = "/Users/lingminjun/work/work_code/utils/src/main/resources/item_list.txt";
        String out = "/Users/lingminjun/work/work_code/utils/src/main/resources/list.xlsx";
        String spl= "\t";

        try {
            ExcelUtils.lineTxtToExcel(in,out,spl);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
