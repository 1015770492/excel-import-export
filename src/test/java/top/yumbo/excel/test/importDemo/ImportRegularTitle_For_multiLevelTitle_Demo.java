package top.yumbo.excel.test.importDemo;

import top.yumbo.excel.test.entity.ImportRegularTitle;
import top.yumbo.excel.test.entity.ImportRegularTitle_For_multiLevelTitle;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.FileInputStream;
import java.util.List;

public class ImportRegularTitle_For_multiLevelTitle_Demo {

    public static void main(String[] args) throws Exception{
        System.out.println("=====导入季度数据======");
        final long start = System.currentTimeMillis();
        String areaQuarter = "src/test/resources/excel/ImportRegularTitle.xls";
        final List<ImportRegularTitle_For_multiLevelTitle> quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(areaQuarter), ImportRegularTitle_For_multiLevelTitle.class, 2000);
        final long end = System.currentTimeMillis();
        System.out.println("总共耗时" + (end - start) + "毫秒");
        quarterList.forEach(System.out::println);
        System.out.println("总共有" + quarterList.size() + "条记录");

    }
}
