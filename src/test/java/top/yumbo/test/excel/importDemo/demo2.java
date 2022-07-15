package top.yumbo.test.excel.importDemo;

import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.FileInputStream;
import java.util.List;

public class demo2 {

    public static void main(String[] args) throws Exception {
        String file = "src/test/java/top/yumbo/test/excel/demo.xls";
        System.out.println("=====投资项目数据======");
        final long start = System.currentTimeMillis();
        List<ImportForInveProj> quarterList;
        try {
            quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(file), 1, ImportForInveProj.class, 2000);
            System.out.println(quarterList);
        } catch (Exception e) {
            e.printStackTrace();
        }
        final long end = System.currentTimeMillis();
        System.out.println("总共耗时" + (end - start) + "毫秒");
    }
}