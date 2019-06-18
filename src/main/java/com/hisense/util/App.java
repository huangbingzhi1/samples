package com.hisense.util;

/**
 * @Author Huang.bingzhi
 * @Date 2019/6/18 17:50
 * @Version 1.0
 */
public class App {
    public static void main(String[] args) throws Exception {
        if (args.length <= -1) {
            System.out.println("输入有误");
        } else {
            String filePath = args[0];
//            String filePath = "E:\\aaa.xlsx";
            ExcelFactory factory = new ExcelFactory(filePath);

            try {
                System.out.println("正在处理文件：" + filePath);
                factory.doBusiness();
                System.out.println("***************处理成功***************");
            } catch (Exception e) {
                System.out.println(e.toString());
            }

        }
    }
}
