package top.yumbo.excel.test.importDemo;

import top.yumbo.excel.test.entity.ImportForYear;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.FileInputStream;
import java.util.List;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
public class ImportForYear_Demo {
    /**
     * 将excel转换为List类型的数据 示例代码
     */
    public static void main(String[] args) throws Exception {

        System.out.println("=====导入年度数据======");
        String areaYear = "src/test/resources/excel/ImportForYear.xlsx";
//        String areaYear = "src/test/resources/excel/ImportForYear_big.xlsx";
        final long start = System.currentTimeMillis();
        final List<ImportForYear> yearList = ExcelImportExportUtils.importExcel(new FileInputStream(areaYear), ImportForYear.class);
        final long end = System.currentTimeMillis();
        System.out.println("总共耗时" + (end - start) + "毫秒");
        System.out.println("总共有" + yearList.size() + "条记录");
        yearList.forEach(System.out::println);

    }
}
