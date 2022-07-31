package top.yumbo.excel.test.entity;

import lombok.Data;
import top.yumbo.excel.annotation.core.ExcelTableHeader;
import top.yumbo.excel.annotation.core.ExcelTitleBind;
import top.yumbo.excel.util.BigDecimalUtils;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
@Data
@ExcelTableHeader(height = 4, sheetName = "区域年度数据", resource = "path:///src/test/java/top/yumbo/test/excel/1.xlsx")
// 表头占4行
public class Export_ImportForYear {


    @ExcelTitleBind(title = "地区", width = 2)
    private String w1;
    @ExcelTitleBind(title = "年份")
    private Integer w2;
    @ExcelTitleBind(title = "地区GDP", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w3;
    @ExcelTitleBind(title = "地区人均GDP", size = BigDecimalUtils.TEN_THOUSAND_STRING)
    private BigDecimal w4;
    @ExcelTitleBind(title = "地区GDP在同级别地区排名")
    private BigDecimal w5;
    @ExcelTitleBind(title = "地区人均GDP在同级别地区排名", size = "0.01")
    private BigDecimal w6;
    @ExcelTitleBind(title = "GDP增速", size = "0.01")
    private BigDecimal w7;
    @ExcelTitleBind(title = "二三产业合计对GDP贡献比例", size = "0.01")
    private BigDecimal w8;
    @ExcelTitleBind(title = "常住人口城镇化率", size = "0.01")
    private BigDecimal w9;
    @ExcelTitleBind(title = "城镇居民人均可支配收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w10;
    @ExcelTitleBind(title = "财政总收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w11;
    @ExcelTitleBind(title = "综合财力", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w12;
    @ExcelTitleBind(title = "一般预算收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w13;
    @ExcelTitleBind(title = "税收收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w14;
    @ExcelTitleBind(title = "非税收收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w15;
    @ExcelTitleBind(title = "政府性基金收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w16;
    @ExcelTitleBind(title = "上级补助收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w17;
    @ExcelTitleBind(title = "返还性收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w18;
    @ExcelTitleBind(title = "一般性转移支付收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w19;
    @ExcelTitleBind(title = "专项转移支付收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w20;
    @ExcelTitleBind(title = "财政总支出", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w21;
    @ExcelTitleBind(title = "一般预算支出", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w22;
    @ExcelTitleBind(title = "一般预算收入在同级别地区排名", size = "0.01")
    private BigDecimal w23;
    @ExcelTitleBind(title = "财政总收入在同级别地区排名", size = "0.01")
    private BigDecimal w24;

    private BigDecimal w25;
    private BigDecimal w26;
    private BigDecimal w27;
    @ExcelTitleBind(title = "上级政府GDP", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w28;
    @ExcelTitleBind(title = "上级政府财政总收入")
    private BigDecimal w29;
    private BigDecimal w30;
    private BigDecimal w31;

}
