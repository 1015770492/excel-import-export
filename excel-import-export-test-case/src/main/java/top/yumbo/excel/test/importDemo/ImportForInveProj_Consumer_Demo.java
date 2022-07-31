package top.yumbo.excel.test.importDemo;

import top.yumbo.excel.test.entity.ImportForInveProj;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.FileInputStream;
import java.util.List;
import java.util.Objects;
import java.util.function.Consumer;

public class ImportForInveProj_Consumer_Demo {

    public static void main(String[] args) throws Exception {
        String file = "src/main/resources/excel/ImportForInveProj.xls";
        System.out.println("=====投资项目数据======");
        final long start = System.currentTimeMillis();
        Consumer<List<ImportForInveProj>> consumer = (quarterList) -> {
            if (Objects.requireNonNull(quarterList).size() > 0) {
                quarterList.forEach(System.out::println);
                System.out.println("总共有" + quarterList.size() + "条记录");
            }
        };
        ExcelImportExportUtils.importExcelConsumer(new FileInputStream(file), ImportForInveProj.class, consumer, 10000);
        final long end = System.currentTimeMillis();
        System.out.println("总共耗时" + (end - start) + "毫秒");

    }
}
