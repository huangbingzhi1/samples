package com.hisense;

import com.hisense.util.ExcelFactory;

/**
 * 启动入口
 */
public class App 
{
    public static void main( String[] args )
    {
        if (args.length <= 0) {
            System.out.println("输入有误");
        } else {
            String filePath = args[0];
            //String filePath = "E:\\满意度.xlsx";
            ExcelFactory factory = null;
            try {
                factory = new ExcelFactory(filePath);
            } catch (Exception e) {
                e.printStackTrace();
            }

            try {
                System.out.println("正在处理文件："+filePath);
                factory.doBusiness();
                System.out.println("处理成功");
            } catch (Exception e) {
                System.out.println(e.toString());
            }

        }
    }
}



