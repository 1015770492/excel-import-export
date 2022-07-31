package top.yumbo.excel.test.exportDemo;

import com.alibaba.fastjson.JSONObject;
import top.yumbo.excel.test.entity.Export_ImportForQuarter;
import top.yumbo.excel.test.entity.ImportForQuarter;
import top.yumbo.excel.entity.TitleBuilder;
import top.yumbo.excel.entity.TitleBuilders;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

/**
 * @author jinhua
 * @date 2021/6/22 21:18
 */
public class ExportSimpleExcelDemo {

    /**
     * 导出简单的Excel文件，不需要模板excel的方式
     */
    public static void main(String[] args) throws Exception {
        // 添加3行表头
        final TitleBuilders titleBuilders = TitleBuilders.builder()
                // 第一行
                .addOneRow(
                        TitleBuilder.builder().title("地区").width(2).height(3).build(),
                        TitleBuilder.builder().title("时间").height(3).build(),
                        TitleBuilder.builder().title("区域债务情况").width(5).build(),
                        TitleBuilder.builder().title("政府可协调性").width(3).build()
                )
                // 第二行
                .addOneRow(
                        TitleBuilder.builder().index(3).title("过去12个月内区域金融新债务违约情况").width(4).build(),
                        TitleBuilder.builder().title("区域偿债统筹管理能力").height(2).build(),
                        TitleBuilder.builder().title("业务合作可协调性").height(2).build(),
                        TitleBuilder.builder().title("还款可协调性").height(2).build(),
                        TitleBuilder.builder().title("数财通系统部署情况").height(2).build()
                )
                // 第三行
                .addOneRow(
                        TitleBuilder.builder().index(3).title("违约主体家数").build(),
                        TitleBuilder.builder().title("合计违约规模").build(),
                        TitleBuilder.builder().title("风险性质").build(),
                        TitleBuilder.builder().title("风险品种").build()
                )
                // 第四行
                .addOneRow(
                        TitleBuilder.builder().title("市州").build(),
                        TitleBuilder.builder().title("区县").build(),
                        TitleBuilder.builder().title("").build(),
                        TitleBuilder.builder().title("个").build(),
                        TitleBuilder.builder().title("亿元").build(),
                        TitleBuilder.builder().title("管理失误违约/技术违约/实质违约").build(),
                        TitleBuilder.builder().title("标准债券/非标集合产品/银行贷款或单一产品").build(),
                        TitleBuilder.builder().title("强/弱").build(),
                        TitleBuilder.builder().title("强/弱").build(),
                        TitleBuilder.builder().title("强/弱").build(),
                        TitleBuilder.builder().title("无部署/部署中/有效部署并应用").build()
                ).build();
        /**
         * 得到List集合
         */
        System.out.println("=====导入季度数据======");
        String areaQuarter = "src/test/resources/excel/ImportForQuarter.xlsx";
//        String areaQuarter = "D:/季度数据-原样式导出6000.xlsx";
        final long start1 = System.currentTimeMillis();
        final List<ImportForQuarter> quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(areaQuarter), ImportForQuarter.class, 30000);
        final long end1 = System.currentTimeMillis();
        System.out.println("数据量" + quarterList.size() + "条，导入耗时" + (end1 - start1) + "毫秒");
        quarterList.forEach(System.out::println);
//        exportHighLight(quarterList, 3000);
        final List<Export_ImportForQuarter> exportForQuarterList = JSONObject.parseArray(JSONObject.toJSONString(quarterList), Export_ImportForQuarter.class);

        ExcelImportExportUtils.exportSimpleExcel(exportForQuarterList,titleBuilders,new FileOutputStream("D:/季度数据-简单导出.xlsx"));
    }

}
