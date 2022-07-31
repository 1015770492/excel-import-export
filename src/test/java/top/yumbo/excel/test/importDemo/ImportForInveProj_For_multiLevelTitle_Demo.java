package top.yumbo.excel.test.importDemo;

import top.yumbo.excel.test.entity.ImportForInveProj_For_multiLevelTitle;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.FileInputStream;
import java.util.List;

public class ImportForInveProj_For_multiLevelTitle_Demo {

    public static void main(String[] args) throws Exception {
        String file = "src/test/resources/excel/ImportForInveProj.xls";
        System.out.println("=====投资项目数据======");
        final long start = System.currentTimeMillis();
        List<ImportForInveProj_For_multiLevelTitle> quarterList;
        try {
            quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(file), ImportForInveProj_For_multiLevelTitle.class, 2000);
            for (int i = 0; i < quarterList.size(); i++) {
                System.out.println("第" + (i + 1) + "行数据：");
                System.out.println(quarterList.get(i));
                System.out.println();
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        final long end = System.currentTimeMillis();
        System.out.println("总共耗时" + (end - start) + "毫秒");
    }
}
