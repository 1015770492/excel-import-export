package top.yumbo.excel.test.entity;

//import io.swagger.annotations.ApiModelProperty;

import lombok.Data;
import top.yumbo.excel.annotation.business.MapEntry;
import top.yumbo.excel.annotation.core.ExcelTableHeader;
import top.yumbo.excel.annotation.core.ExcelTitleBind;

/**
 * 投资项目导入类
 * @author: jinhua Yu
 * @createDate:2022/7/14
 * @description:
 */
// 默认globalTitleSplit就是下划线
@Data
@ExcelTableHeader(height = 8, sheetName = "项目总表",globalTitleSplit = "_")
public class ImportForInveProj_For_multiLevelTitle {

    // 1.下达年份批次  默认titleSplit 使用的是下划线，为了不写死，因为可能有些标题本身带下划线，可以换成@或者其他字符串
    @ExcelTitleBind(title = "中央预算内投资资金情况_下达年份批次",titleSplit = "_")
    private String w3;
    // 1.下达资金
    @ExcelTitleBind(title = "中央预算内投资资金情况_下达资金")
    private String w4;
    // 1.支付资金
    @ExcelTitleBind(title = "中央预算内投资资金情况_支付资金")
    private String w5;
    // 1.支付进度
    @ExcelTitleBind(title = "中央预算内投资资金情况_支付进度")
    private String w6;

    // 2.下达年份批次
    @ExcelTitleBind(title = "省预算内投资资金情况_下达年份批次")
    private String w7;
    @ExcelTitleBind(title = "省预算内投资资金情况_资金类型")
    private String w8;
    // 2.下达资金（万元）
    @ExcelTitleBind(title = "省预算内投资资金情况_下达资金（万元）")
    private String w9;
    // 2.支付资金（万元）
    @ExcelTitleBind(title = "省预算内投资资金情况_支付资金（万元）")
    private String w10;
    // 2.支付进度
    @ExcelTitleBind(title = "省预算内投资资金情况_支付进度")
    private String w11;


    // 3.下达年份批次
    @ExcelTitleBind(title = "来地方政府专项债券资金情况_下达年份批次")
    private String w12;
    // 3.下达资金（万元）
    @ExcelTitleBind(title = "来地方政府专项债券资金情况_下达资金（万元）")
    private String w13;
    // 3.支付资金（万元）
    @ExcelTitleBind(title = "来地方政府专项债券资金情况_支付资金（万元）")
    private String w14;
    // 3.支付进度
    @ExcelTitleBind(title = "来地方政府专项债券资金情况_支付进度")
    private String w15;


}
