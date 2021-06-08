package top.yumbo.test.excel.importDemo;

import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.FileInputStream;
import java.util.List;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
public class ImportExcelDemo {
    /**
     * 将excel转换为List类型的数据 示例代码
     */
    public static void main(String[] args) throws Exception{

        System.out.println("=====导入年度数据======");
        String areaYear = "src/test/java/top/yumbo/test/excel/1.xlsx";
        final List<ImportForYear> yearList = ExcelImportExportUtils.importExcel(new FileInputStream(areaYear), ImportForYear.class, "xlsx");
        yearList.forEach(System.out::println);
        System.out.println("=====导入季度数据======");
        String areaQuarter = "src/test/java/top/yumbo/test/excel/2.xlsx";
        final List<ImportForQuarter> quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(areaQuarter), ImportForQuarter.class, "xlsx");
        quarterList.forEach(System.out::println);
    }
}
