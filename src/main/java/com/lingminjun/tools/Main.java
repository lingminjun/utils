package com.lingminjun.tools;

import com.lingminjun.tools.excel.ExcelUtils;

import java.io.IOException;

/**
 * Created by lingminjun on 16/12/15.
 */
public class Main {
    private static void help() {
        System.out.println(
                "--txt_to_excel IN_PUT_PATH OUT_PUT_PATH [SPLIT_STRING] 按照行读取文本并转换成excel\n" +
                        "\tIN_PUT_PATH: 输入文本路径,不能为空\n" +
                        "\tOUT_PUT_PATH: 输出excel地址,不能为空,请自带后缀名xlsx\n" +
                        "\tSPLIT_STRING: 分割字符,不传时默认为空格分割\n" +
                        "\n" +
                "--split_excel IN_PUT_PATH OUT_PUT_PATH MAX_COUNT 将一个excel进行拆分\n" +
                        "\tIN_PUT_PATH: 输入文本路径,不能为空\n" +
                        "\tOUT_PUT_PATH: 输出excel地址,不能为空,请自带后缀名xlsx\n" +
                        "\tMAX_COUNT: 分割后excel最大行数控制\n" +
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
        String out = "/Users/lingminjun/work/work_code/utils/src/main/resources/test.xls";
        String out1 = "/Users/lingminjun/work/work_code/utils/src/main/resources/result/list.xlsx";
        String spl= "\t";

        try {
//            ExcelUtils.lineTxtToExcel(in,out,spl);
            ExcelUtils.splitExcel(out,out1,6000);
        } catch (IOException e) {
            e.printStackTrace();
        }

//        Other other = new Other();
//        Map map = other.getUisa();
//        System.out.println(map.toString());
//        other.pares("[{\"type\":\"REDUCE\",\"name\":\"满减送\",\"rgb\":\"0xf13c34\"},{\"type\":\"DISCOUNT\",\"name\":\"满件折\",\"rgb\":\"0x6092f3\"},{\"type\":\"FLASH\",\"name\":\"限时促销\",\"rgb\":\"0xf4ae35\"},{\"type\":\"MIX_DISCOUNT\",\"name\":\"搭配折扣\",\"rgb\":\"0xe9378b\"},{\"type\":\"NYRX\",\"name\":\"N元任选\",\"rgb\":\"0x35b79b\"},{\"type\":\"GIFT\",\"name\":\"换购\",\"rgb\":\"0x9974ff\"}]");
//        System.out.println(map.toString());
    }
}
