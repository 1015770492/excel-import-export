package top.yumbo.test.excel.exportDemo;

import lombok.Data;
import top.yumbo.excel.annotation.core.TitleBind;
import top.yumbo.excel.annotation.core.TableHeader;
import top.yumbo.excel.util.BigDecimalUtils;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
@Data
@TableHeader(height = 4, tableName = "区域年度数据", resource = "path:///src/test/java/top/yumbo/test/excel/1.xlsx")
// 表头占4行
public class ExportForYear {


    @TitleBind(title = "地区", width = 2)
    private String w1;
    @TitleBind(title = "年份")
    private Integer w2;
    @TitleBind(title = "地区GDP", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w3;
    @TitleBind(title = "地区人均GDP", size = BigDecimalUtils.TEN_THOUSAND_STRING)
    private BigDecimal w4;
    @TitleBind(title = "地区GDP在同级别地区排名")
    private BigDecimal w5;
    @TitleBind(title = "地区人均GDP在同级别地区排名", size = "0.01")
    private BigDecimal w6;
    @TitleBind(title = "GDP增速", size = "0.01")
    private BigDecimal w7;
    @TitleBind(title = "二三产业合计对GDP贡献比例", size = "0.01")
    private BigDecimal w8;
    @TitleBind(title = "常住人口城镇化率", size = "0.01")
    private BigDecimal w9;
    @TitleBind(title = "城镇居民人均可支配收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w10;
    @TitleBind(title = "财政总收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w11;
    @TitleBind(title = "综合财力", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w12;
    @TitleBind(title = "一般预算收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w13;
    @TitleBind(title = "税收收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w14;
    @TitleBind(title = "非税收收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w15;
    @TitleBind(title = "政府性基金收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w16;
    @TitleBind(title = "上级补助收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w17;
    @TitleBind(title = "返还性收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w18;
    @TitleBind(title = "一般性转移支付收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w19;
    @TitleBind(title = "专项转移支付收入", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w20;
    @TitleBind(title = "财政总支出", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w21;
    @TitleBind(title = "一般预算支出", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w22;
    @TitleBind(title = "一般预算收入在同级别地区排名", size = "0.01")
    private BigDecimal w23;
    @TitleBind(title = "财政总收入在同级别地区排名", size = "0.01")
    private BigDecimal w24;

    private BigDecimal w25;
    private BigDecimal w26;
    private BigDecimal w27;
    @TitleBind(title = "上级政府GDP", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal w28;
    @TitleBind(title = "上级政府财政总收入")
    private BigDecimal w29;
    private BigDecimal w30;
    private BigDecimal w31;

}
