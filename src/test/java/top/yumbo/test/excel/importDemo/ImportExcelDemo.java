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
//        String areaYear = "src/test/java/top/yumbo/test/excel/1.xlsx";
        String areaYear = "src/test/java/top/yumbo/test/excel/1_big.xlsx";
        final long start = System.currentTimeMillis();
        final List<ImportForYear> yearList = ExcelImportExportUtils.importExcel(new FileInputStream(areaYear), ImportForYear.class);
        final long end = System.currentTimeMillis();
        System.out.println("总共耗时"+(end-start)+"毫秒");
        System.out.println("总共有"+yearList.size()+"条记录");
//        yearList.forEach(System.out::println);

        System.out.println("=====导入季度数据======");
        final long start2 = System.currentTimeMillis();
//        String areaQuarter = "src/test/java/top/yumbo/test/excel/2.xlsx";
        String areaQuarter = "src/test/java/top/yumbo/test/excel/2_big.xlsx";
        final List<ImportForQuarter> quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(areaQuarter), ImportForQuarter.class,2000);
        final long end2 = System.currentTimeMillis();
        System.out.println("总共耗时"+(end-start)+"毫秒");
//        quarterList.forEach(System.out::println);
        System.out.println("总共有"+quarterList.size()+"条记录");
    }
}
