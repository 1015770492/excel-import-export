package top.yumbo.test.excel.exportDemo;

import com.alibaba.fastjson.JSONArray;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import top.yumbo.excel.entity.CellStyleEntity;
import top.yumbo.excel.entity.TitleCellStylePredicate;
import top.yumbo.excel.entity.TitlePredicateList;
import top.yumbo.excel.util.ExcelImportExportUtils;
import top.yumbo.test.excel.importDemo.ExcelImportTemplateForQuarter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
public class ExcelExportDemo {

    interface PredicateCellStyle<T> {
        default boolean enable(T t, Predicate predicate) {
            return predicate.test(t);
        }
    }

    public static void main(String[] args) {
        System.out.println("=====季度数据======");
        String fileName = "2.xlsx";
        List<ExcelImportTemplateForQuarter> list = importData(fileName, ExcelImportTemplateForQuarter.class);
        System.out.println(list);
        final List<ExcelExportTemplateForQuarter> quarterList = JSONArray.parseArray(JSONArray.toJSONString(list), ExcelExportTemplateForQuarter.class);


        final String concurrentPath = getConcurrentPath();
        String relativePath = "/src/test/java/top/yumbo/test/excel";
        try (FileInputStream fis = new FileInputStream(concurrentPath + relativePath + "/test2.xlsx");) {

            Workbook workbook = new XSSFWorkbook(fis);
            final Sheet sheet = workbook.getSheetAt(0);
//            ExcelImportExportUtils.filledListToSheet(quarterList, workbook.getSheetAt(0));
//            ExcelImportExportUtils.filledListToSheetWithCellStyle(quarterList, workbook.getSheetAt(0));
//            /**
//             * 指定某一标题列进行高亮展示 并且条件过滤某些行 高亮显示
//             */
//            final CellStyle cellStyle = CellStyleEntity.builder().fontName("微软雅黑").fontSize(12).bgColor(9).build().getCellStyle(workbook);
//            ExcelImportExportUtils.filledListToSheetWithCellStyleByFieldPredicate(quarterList, cellStyle,ExcelExportTemplateForQuarter::getBreachTotalScale, (one) -> {
//                if (one.getQuarter() > 3) {
//                    return true;
//                }
//                return false;
//            }, workbook.getSheetAt(0));

            final CellStyle cellStyle = CellStyleEntity.builder().fontName("微软雅黑").bold(true).fontSize(12).build().getCellStyle(workbook);
            final TitlePredicateList<ExcelExportTemplateForQuarter> predicateList = new TitlePredicateList<>();
            // 提供断言处理
            Predicate<ExcelExportTemplateForQuarter> predicate = (e) -> {
                String regex = ".*市";// 高亮市
                final Pattern pattern = Pattern.compile(regex);
                final Matcher matcher = pattern.matcher(e.getRegionCode().split(",")[0]);
                if (matcher.matches()) {
                    return true;
                }
                return false;
            };
            Predicate<ExcelExportTemplateForQuarter> predicate2 = (e) -> {
                String regex = ".*市";// 高亮市
                final Pattern pattern = Pattern.compile(regex);
                final Matcher matcher = pattern.matcher(e.getRegionCode().split(",")[1]);
                if (matcher.matches()) {
                    return true;
                }
                return false;
            };
            final List<TitleCellStylePredicate<ExcelExportTemplateForQuarter>> titlePredicateList = predicateList
//                    .add("市州", cellStyle, predicate)
//                    .add("区县", cellStyle, predicate2)
                    .getTitlePredicateList();
            ExcelImportExportUtils.filledListToSheetWithCellStyleByBatchTitlePredicate(quarterList, titlePredicateList, sheet);

//            /**
//             * 指定某一标题列进行高亮展示 并且条件过滤某些行 高亮显示
//             */
//            final CellStyle cellStyle = CellStyleEntity.builder().fontName("微软雅黑").bold(true).fontSize(12).build().getCellStyle(workbook);
//            ExcelImportExportUtils.filledListToSheetWithCellStyleByTitlePredicate(quarterList, cellStyle, "区县", (one) -> {
//                String regex = ".*市";// 高亮市
//                final Pattern pattern = Pattern.compile(regex);
//                final Matcher matcher = pattern.matcher(one.getRegionCode());
//                if (matcher.matches()) {
//                    return true;
//                }
//                return false;
//            }, workbook.getSheetAt(0));


//            /**
//             * 某些行高亮展示
//             */
//            final List<CellStyle> cellStyleList = Arrays.asList(
//                    CellStyleEntity.builder().fontName("微软雅黑").fontSize(12).bgColor(9).build().getCellStyle(workbook),
//                    CellStyleEntity.builder().fontName("宋体").fontSize(12).bgColor(9).foregroundColor(13).build().getCellStyle(workbook),
//                    CellStyleEntity.builder().fontName("微软雅黑").fontSize(12).bgColor(10).build().getCellStyle(workbook)
//            );
//            ExcelImportExportUtils.filledListToSheetWithCellStyleByFunction(quarterList, cellStyleList, (one) -> {
//                if (one.getYear() % 2 == 0 || one.getQuarter() == 3) {
//                    return 1;
//                }else {
//                    return 0;
//                }
//            }, workbook.getSheetAt(0));

            workbook.createSheet("证明是新的");
            workbook.write(new FileOutputStream("D:/导出的季度数据.xlsx"));
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getConcurrentPath() {
        File directory = new File("");//设定为当前文件夹
        String currentAbsolutePath = directory.getAbsolutePath();
        return currentAbsolutePath;
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
}
