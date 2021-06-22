package top.yumbo.test.excel.exportDemo;

import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import top.yumbo.excel.entity.CellStyleBuilder;
import top.yumbo.excel.util.ExcelImportExportUtils;
import top.yumbo.test.excel.importDemo.ImportForQuarter;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.List;

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
//        String areaQuarter = "D:/季度数据-原样式导出6000.xlsx";
        final long start1 = System.currentTimeMillis();
        final List<ImportForQuarter> quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(areaQuarter), ImportForQuarter.class, 30000);
        final long end1 = System.currentTimeMillis();
        System.out.println("数据量" + quarterList.size() + "条，导入耗时" + (end1 - start1) + "毫秒");
        quarterList.forEach(System.out::println);
//        exportHighLight(quarterList, 3000);
        final List<ExportForQuarter> exportForQuarterList = JSONObject.parseArray(JSONObject.toJSONString(quarterList), ExportForQuarter.class);
        exportDefault(exportForQuarterList, 3000);

        /**
         * 将其导出
         */
//        if (quarterList != null) {
//            // 将数据导出到本地文件, 如果要导出到web暴露出去只要传入输出流即可
//            List<ExportForQuarter> list = new ArrayList<>();
//            for (int i = 0; i < 5; i++) {
//                list.addAll(exportForQuarterList);
//            }
//            for (int i = 0; i < 3; i++) {
//                System.out.println("第" + (i + 1) + "次导出测试");
//                System.out.println("总数据量：" + list.size() + "条记录");
//                IntStream.of(10000).forEach(threshold -> {
//                    System.out.println("threshold=" + threshold);
//                    try {
//                        exportDefault(list, threshold);
//                        exportHighLight(list, threshold);
//                        System.out.println(">>>>>>>>>>>>>>>>>");
//
//                    } catch (Exception e) {
//                        e.printStackTrace();
//                    }
//                });
//                break;
//            }
//        }

    }

    private static void exportHighLight(List<ExportForQuarter> quarterList, int threshold) throws Exception {
        /**
         * 高亮行
         */
        final long start = System.currentTimeMillis();
        ExcelImportExportUtils.exportExcelRowHighLight(quarterList,
                new FileOutputStream("D:/季度数据-高亮行导出" + threshold + ".xlsx"),
                (t) -> {
                    if (t.getW2() == 1) {
                        return IndexedColors.YELLOW;
                    } else if (t.getW2() == 2) {
                        return IndexedColors.ROSE;
                    } else if (t.getW2() == 3) {
                        return IndexedColors.SKY_BLUE;
                    } else if (t.getW2() == 4) {
                        return IndexedColors.GREY_25_PERCENT;
                    } else {
                        return IndexedColors.WHITE;
                    }
                }, threshold);
        final long end = System.currentTimeMillis();
        System.out.println("高亮行总共用了" + (end - start) + "毫秒\n");
    }

    private static int exportDefault(List<ExportForQuarter> quarterList, int threshold) throws Exception {
        /**
         * 原样式导出
         */
        final long start1 = System.currentTimeMillis();
        ExcelImportExportUtils.exportExcel(quarterList, new FileOutputStream("D:/季度数据-原样式导出" + threshold + ".xlsx"), threshold);
        final long end1 = System.currentTimeMillis();
        System.out.println("原样式导出总共用了" + (end1 - start1) + "毫秒\n");
        return threshold;
    }

    /**
     * 高亮行（断言方式高亮示例代码）
     * 高亮符合条件的行
     */
    private static void rowHighLight(List<ExportForQuarter> quarterList) throws Exception {
        quarterList.forEach(System.out::println);
        /**
         * 某些行高亮展示，字体等其他样式继续进行链式调用即可设置
         */
        // 3种样式
        final List<CellStyle> cellStyleList = Arrays.asList(
                CellStyleBuilder.builder().foregroundColor(51).fontName("微软雅黑").build().getCellStyle(),// 灰色
                CellStyleBuilder.builder().foregroundColor(12).build().getCellStyle(),// 蓝色
                CellStyleBuilder.builder().foregroundColor(13).build().getCellStyle(),// 黄色
                CellStyleBuilder.builder().foregroundColor(17).build().getCellStyle(),// 绿色
                CellStyleBuilder.builder().build().getCellStyle()// 绿色
        );

    }



}
