package top.yumbo.test.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.yumbo.excel.util.ExcelImportExportUtils;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;


public class ExcelImportTest {
    /**
     * 将excel转换为List类型的数据 示例代码
     */
    public static void main(String[] args) {
        System.out.println("=====年度数据======");
        String fileName = "1.xlsx";
        importData(fileName, RegionYearETLSyncResponse.class);
        System.out.println("=====季度数据======");
        fileName = "2.xlsx";
        importData(fileName, RegionQuarterETLSyncResponse.class);
    }

    private static void importData(String fileName, Class clazz) {
        String type = fileName.split("\\.")[1];
        System.out.println("=======");
        File directory = new File("");//设定为当前文件夹
        String currentAbsolutePath = directory.getAbsolutePath();
        //System.out.println(currentAbsolutePath);


        String relativePath = "src/test/java/top/yumbo/test/excel/" + fileName;

        try (FileInputStream fis = new FileInputStream(currentAbsolutePath + "/" + relativePath);) {
            Workbook sheets = null;
            if ("xls".equals(type)) {
                sheets = new HSSFWorkbook(fis);
            } else if ("xlsx".equals(type)) {
                sheets = new XSSFWorkbook(fis);
            }

            if (sheets != null) {
                Sheet sheet = sheets.getSheetAt(0);
                /**
                 * 核心方法，传入泛型（带注解信息），sheet待解析的数据
                 */
                // 加了注解信息的实体类
                List list = ExcelImportExportUtils.parseSheetToList(clazz, sheet);
                list.forEach(System.out::println);

            }

        } catch (Exception e) {

        }
    }
}
