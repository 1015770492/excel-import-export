package top.yumbo.test.excel.exportDemo;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import top.yumbo.excel.entity.CellStyleEntity;
import top.yumbo.excel.entity.TitleCellStylePredicate;
import top.yumbo.excel.entity.TitlePredicateList;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.List;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
public class ExportExcelDemo {

    public static void main(String[] args) throws Exception {
        /**
         * 得到List集合
         */
        System.out.println("=====导入季度数据======");
        String areaQuarter = "src/test/java/top/yumbo/test/excel/2.xlsx";
        final List<ExportForQuarter> quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(areaQuarter), ExportForQuarter.class, "xlsx");

        /**
         * 将其导出
         */
        if (quarterList != null) {
            quarterList.forEach(System.out::println);
            // 将数据导出到本地文件,如果要导出到web暴露出去只要传入输出流即可
            ExcelImportExportUtils.exportExcel(quarterList, new FileOutputStream("D:/季度数据-原样式导出.xlsx"));
        }

    }

    /**
     * 高亮符合条件的行
     */
    private static void rowHighLight(List<ExportForQuarter> quarterList, Workbook workbook) throws Exception {
        /**
         * 某些行高亮展示
         */
        // 3种样式
        final List<CellStyle> cellStyleList = Arrays.asList(
                CellStyleEntity.builder().fontName("微软雅黑").fontSize(12).bgColor(9).build().getCellStyle(workbook),
                CellStyleEntity.builder().fontSize(12).bgColor(9).foregroundColor(13).build().getCellStyle(workbook),
                CellStyleEntity.builder().fontName("微软雅黑").fontSize(12).bgColor(10).build().getCellStyle(workbook)
        );
        final Function<ExportForQuarter, Integer> functional = (one) -> {
            if (one.getRiskNature().equals("技术违约") && (one.getYear() % 2 == 0 || one.getQuarter() == 3)) {
                return 1;
            } else {
                return 0;
            }
        };
        // 使用函数时接口返回样式的下标，然后就会将样式注入进去
        ExcelImportExportUtils.filledListToSheetWithCellStyleByFunction(quarterList, cellStyleList, (one) -> {
            if (one.getRiskNature().equals("技术违约") && (one.getYear() % 2 == 0 || one.getQuarter() == 3)) {
                return 1;
            } else {
                return 0;
            }
        }, workbook.getSheetAt(0));
    }

    /**
     * 指定标题下的单元格 部分高亮
     */
    private static void titlePredicate(List<ExportForQuarter> quarterList, Workbook workbook, Sheet sheet) throws Exception {
        final CellStyle cellStyle = CellStyleEntity.builder().fontName("微软雅黑").bold(true).fontSize(12).build().getCellStyle(workbook);
        //
        final CellStyle cellStyle3 = CellStyleEntity.builder().fontSize(12).fontColor(14).foregroundColor(13).build().getCellStyle(workbook);
        final CellStyle cellStyle4 = CellStyleEntity.builder().fontSize(12).fontColor(10).bold(true).fontColor(14).foregroundColor(40).build().getCellStyle(workbook);
        final TitlePredicateList<ExportForQuarter> predicateList = new TitlePredicateList<>();
        // 提供断言处理
        Predicate<ExportForQuarter> predicate = (e) -> {
            String regex = ".*市";// 高亮市
            final Pattern pattern = Pattern.compile(regex);
            final Matcher matcher = pattern.matcher(e.getRegionCode().split(",")[0]);
            if (matcher.matches()) {
                return true;
            }
            return false;
        };
        Predicate<ExportForQuarter> predicate2 = (e) -> {
            String regex = ".*市";// 高亮市
            final Pattern pattern = Pattern.compile(regex);
            final Matcher matcher = pattern.matcher(e.getRegionCode().split(",")[1]);
            if (matcher.matches()) {
                return true;
            }
            return false;
        };
        Predicate<ExportForQuarter> predicate3 = (e) -> {
            if (e.getRiskNature().equals("管理失误违约")) {
                return true;
            }
            return false;
        };
        // 高亮时间，第3季度的背景色设置为蓝色，字体红色加粗
        Predicate<ExportForQuarter> predicate4 = (e) -> {
            if (e.getQuarter() == 3) {
                return true;
            }
            return false;
        };

        final List<TitleCellStylePredicate<ExportForQuarter>> titlePredicateList = predicateList
                .add("市州", cellStyle, predicate)
                .add("区县", cellStyle, predicate2)
                .add("风险性质", cellStyle3, predicate3)
                .add("时间", cellStyle4, predicate4)
                .getTitlePredicateList();
        ExcelImportExportUtils.filledListToSheetWithCellStyleByBatchTitlePredicate(quarterList, titlePredicateList, sheet);
    }

    private static String getConcurrentPath() {
        File directory = new File("");//设定为当前文件夹
        String currentAbsolutePath = directory.getAbsolutePath();
        return currentAbsolutePath;
    }


}
