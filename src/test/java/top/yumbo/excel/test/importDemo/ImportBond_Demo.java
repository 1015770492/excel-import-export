package top.yumbo.excel.test.importDemo;

import top.yumbo.excel.test.entity.BIExcelResp;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.FileInputStream;
import java.util.List;

/**
 * @author jinhua
 * @date 2021/6/18 17:23
 */
public class ImportBond_Demo {


    public static void main(String[] args) throws Exception{
        System.out.println("=====导入年度数据======");
        String areaYear = "src/test/resources/excel/BIExcelResp.xlsx";
        final long start = System.currentTimeMillis();
        final List<BIExcelResp> bList = ExcelImportExportUtils.importExcel(new FileInputStream(areaYear), BIExcelResp.class,2000);
        final long end = System.currentTimeMillis();
        System.out.println("总共耗时"+(end-start)+"毫秒");
        System.out.println("总共有"+bList.size()+"条记录");
        bList.forEach(System.out::println);

    }

}
