package top.yumbo.excel.test.importDemo;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import top.yumbo.excel.test.entity.ImportPIM;
import top.yumbo.excel.util.ExcelImportUtils2;

import java.io.FileInputStream;
import java.util.List;

/**
 * @author jinhua
 * @date 2021/7/1 0:11
 */
public class ImportPIM_Demo {
    public static void main(String[] args) throws Exception{
        String xlsx = "src/main/resources/excel/ImportPIM.xlsx";

        final List<ImportPIM> pimExcels = ExcelImportUtils2.importExcel(WorkbookFactory.create(new FileInputStream(xlsx)).getSheetAt(0), ImportPIM.class);
        pimExcels.forEach(System.out::println);
    }
}
