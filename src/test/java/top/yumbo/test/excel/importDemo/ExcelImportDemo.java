package top.yumbo.test.excel.importDemo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.*;
import java.util.List;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
public class ExcelImportDemo {
    /**
     * 将excel转换为List类型的数据 示例代码
     */
    public static void main(String[] args) {
        System.out.println("=====导入年度数据======");
        String fileName = "1.xlsx";
        importData(fileName, ExcelImportTemplateForYear.class);
        System.out.println("=====导出季度数据======");
        fileName = "2.xlsx";
        importData(fileName, ExcelImportTemplateForQuarter.class);
    }

    private static <T> List<T> importData(String fileName, Class<T> clazz) {
        String type = fileName.split("\\.")[1];
        String currentAbsolutePath = getConcurrentPath();
        //System.out.println(currentAbsolutePath);

        List<T> list = null;

        String relativePath = "src/test/java/top/yumbo/test/excel/" + fileName;

        try (FileInputStream fis = new FileInputStream(currentAbsolutePath + "/" + relativePath);) {
            Workbook workbook = null;
            if ("xls".equals(type)) {
                workbook = new HSSFWorkbook(fis);
            } else if ("xlsx".equals(type)) {
                workbook = new XSSFWorkbook(fis);
            }

            if (workbook != null) {
                Sheet sheet = workbook.getSheetAt(0);

                /**
                 * 核心方法，传入泛型（带注解信息），sheet待解析的数据
                 */
                // 加了注解信息的实体类
                list = ExcelImportExportUtils.parseSheetToList(clazz, sheet);
                list.forEach(System.out::println);

            }

        } catch (Exception e) {

        }
        return list;
    }

    private static String getConcurrentPath() {
        File directory = new File("");//设定为当前文件夹
        String currentAbsolutePath = directory.getAbsolutePath();
        return currentAbsolutePath;
    }
}
