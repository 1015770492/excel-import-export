package top.yumbo.test.excel.exportDemo;

import lombok.Data;
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.util.BigDecimalUtils;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
@Data
@ExcelTableHeader(height = 4, sheetName = "区域年度数据", resource = "path:///src/test/java/top/yumbo/test/excel/1.xlsx")
// 表头占4行
public class ExportForYear {


    @ExcelCellBind(title = "地区", width = 2)
    private String w1;
    @ExcelCellBind(title = "年份")
    private Integer w2;
    @ExcelCellBind(title = "地区GDP", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w3;
    @ExcelCellBind(title = "地区人均GDP", size = BigDecimalUtils.TEN_THOUSAND_STRING)
    private BigDecimal w4;
    @ExcelCellBind(title = "地区GDP在同级别地区排名")
    private BigDecimal w5;
    @ExcelCellBind(title = "地区人均GDP在同级别地区排名", size = "0.01")
    private BigDecimal w6;
    @ExcelCellBind(title = "GDP增速", size = "0.01")
    private BigDecimal w7;
    @ExcelCellBind(title = "二三产业合计对GDP贡献比例", size = "0.01")
    private BigDecimal w8;
    @ExcelCellBind(title = "常住人口城镇化率", size = "0.01")
    private BigDecimal w9;
    @ExcelCellBind(title = "城镇居民人均可支配收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w10;
    @ExcelCellBind(title = "财政总收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w11;
    @ExcelCellBind(title = "综合财力", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w12;
    @ExcelCellBind(title = "一般预算收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w13;
    @ExcelCellBind(title = "税收收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w14;
    @ExcelCellBind(title = "非税收收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w15;
    @ExcelCellBind(title = "政府性基金收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w16;
    @ExcelCellBind(title = "上级补助收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w17;
    @ExcelCellBind(title = "返还性收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w18;
    @ExcelCellBind(title = "一般性转移支付收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w19;
    @ExcelCellBind(title = "专项转移支付收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w20;
    @ExcelCellBind(title = "财政总支出", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w21;
    @ExcelCellBind(title = "一般预算支出", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w22;
    @ExcelCellBind(title = "一般预算收入在同级别地区排名", size = "0.01")
    private BigDecimal w23;
    @ExcelCellBind(title = "财政总收入在同级别地区排名", size = "0.01")
    private BigDecimal w24;

    private BigDecimal w25;
    private BigDecimal w26;
    private BigDecimal w27;
    @ExcelCellBind(title = "上级政府GDP", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w28;
    @ExcelCellBind(title = "上级政府财政总收入")
    private BigDecimal w29;
    private BigDecimal w30;
    private BigDecimal w31;

}
