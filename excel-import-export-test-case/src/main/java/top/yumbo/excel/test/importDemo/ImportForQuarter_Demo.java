package top.yumbo.excel.test.importDemo;

import top.yumbo.excel.test.entity.ImportForQuarter;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.FileInputStream;
import java.util.List;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
public class ImportForQuarter_Demo {
    /**
     * 将excel转换为List类型的数据 示例代码
     */
    public static void main(String[] args) throws Exception {


        System.out.println("=====导入季度数据======");
        final long start = System.currentTimeMillis();
        String areaQuarter = "src/main/resources/excel/ImportForQuarter.xlsx";
//        String areaQuarter = "src/main/resources/excel/ImportForQuarter_big.xlsx";
        final List<ImportForQuarter> quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(areaQuarter), ImportForQuarter.class, 2000);
        final long end = System.currentTimeMillis();
        System.out.println("总共耗时" + (end - start) + "毫秒");
        quarterList.forEach(System.out::println);
        System.out.println("总共有" + quarterList.size() + "条记录");


    }
}
