package top.yumbo.test.excel.importDemo;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import top.yumbo.excel.util.ExcelAnnotationImport;

import java.io.FileInputStream;
import java.util.List;

/**
 * @author jinhua
 * @date 2021/7/1 0:11
 */
public class PIMImportDemo {
    public static void main(String[] args) throws Exception{
        String xlsx = "src/test/java/top/yumbo/test/excel/importDemo/PIM.xlsx";

        final List<PIMExcel> pimExcels = ExcelAnnotationImport.importExcel(WorkbookFactory.create(new FileInputStream(xlsx)).getSheetAt(0), PIMExcel.class);
        pimExcels.forEach(System.out::println);
    }
}
