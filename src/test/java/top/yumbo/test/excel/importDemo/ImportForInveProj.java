package top.yumbo.test.excel.importDemo;

//import io.swagger.annotations.ApiModelProperty;

import lombok.Data;
import top.yumbo.excel.annotation.business.MapEntry;
import top.yumbo.excel.annotation.core.ExcelTableHeader;
import top.yumbo.excel.annotation.core.ExcelTitleBind;

import java.math.BigDecimal;
import java.util.Date;

/**
 * 投资项目导入类
 * @author: Haiming Yu
 * @createDate:2022/7/14
 * @description:
 */
@Data
@ExcelTableHeader(height = 8, sheetName = "项目总表")
public class ImportForInveProj {

    /** 项目名称 */
    @ExcelTitleBind(title = "项目名称")
    private String projectName;

    /** 审批监管平台代码 */
    @ExcelTitleBind(title = "审批监管平台代码")
    private String supervisionCode;

    /** 建设地点 */
    @ExcelTitleBind(title = "建设地点")
    private String constructionAddress;

    /** 行业主管部门 */
    @ExcelTitleBind(title = "行业主管部门")
    private String manageDepartment;

//    /** 所属批次 */
//    @ExcelTitleBind(title = "所属批次")
//    private String belongBatch;

    /** 建设内容及规模 */
    @ExcelTitleBind(title = "建设内容及规模（50字以内）")
    private String constructionContent;

    /** 总投资 （万元） */
    @ExcelTitleBind(title = "总投资\n（万元）",size = "10000")
    private BigDecimal totalInvestmentAmount;

    /**累计总投资 （万元）*/
    @ExcelTitleBind(title = "累计完成投资\n（万元）",size = "10000")
    private BigDecimal w1;

    /** 2022年计划完成投资 （万元） */
    @ExcelTitleBind(title = "2022年计划完成投资\n（万元）")
    private String thisYearPlanInvestmentAmount;

    /** 开工时间 */
    @ExcelTitleBind(title = "开工时间\n20XX.0X")
    private String constructionStartTime;

    /** 完工时间 */
    @ExcelTitleBind(title = "完工时间\n20XX.0X")
    private String finishTime;

    /** 项目推进情况 */
    @ExcelTitleBind(title = "项目推进情况（50字以内）")
    private String progressDetail;

    @ExcelTitleBind(title = "是否开工")
    private String w2;

    /** 是否已入库 */
    @ExcelTitleBind(title = "是否已入库")
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    private String regulated;

    /** 是否申请中央预算内投资 */
    @ExcelTitleBind(title = "是否申请中央预算内投资",nullable = true)
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    private String applyCentralFund;

    /** 是否申请省预算内投资 */
    @ExcelTitleBind(title = "是否申请省预算内投资",nullable = true)
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    private String applyProvinceFund;

    /** 是否申请地方政府专项债券 */
    @ExcelTitleBind(title = "是否申请地方政府专项债券")
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    private String applyLocalGovFund;


    // 1.下达年份批次
    @ExcelTitleBind(title = "中央预算内投资资金情况",nullable = true)
    private String w3;
    // 1.下达资金
    @ExcelTitleBind(positionTitle = "中央预算内投资资金情况",offset = 1,nullable = true)
    private String w4;
    // 1.支付资金
    @ExcelTitleBind(positionTitle = "中央预算内投资资金情况",offset = 2,nullable = true)
    private String w5;
    // 1.支付进度
    @ExcelTitleBind(positionTitle = "中央预算内投资资金情况",offset = 3,nullable = true)
    private String w6;

    // 2.下达年份批次
    @ExcelTitleBind(title = "省预算内投资资金情况",nullable = true)
    private String w7;
    // 资金类型 ：该标题在标题头中唯一，可以直接用title,也可以使用注释掉的注解获取，取决于你的个人理解（用那种都随意）
    @ExcelTitleBind(title = "资金类型",nullable = true)
    //@ExcelTitleBind(positionTitle = "省预算内投资资金情况",offset = 1,nullable = true)
    private String w8;
    // 2.下达资金（万元）
    @ExcelTitleBind(positionTitle = "省预算内投资资金情况",offset = 2,nullable = true)
    private String w9;
    // 2.支付资金（万元）
    @ExcelTitleBind(positionTitle = "省预算内投资资金情况",offset = 3,nullable = true)
    private String w10;
    // 2.支付进度
    @ExcelTitleBind(positionTitle = "省预算内投资资金情况",offset = 4,nullable = true)
    private String w11;


    // 3.下达年份批次
    @ExcelTitleBind(title = "来地方政府专项债券资金情况",nullable = true)
    private String w12;
    // 3.下达资金（万元）
    @ExcelTitleBind(positionTitle = "来地方政府专项债券资金情况",offset = 1,nullable = true)
    private String w13;
    // 3.支付资金（万元）
    @ExcelTitleBind(positionTitle = "来地方政府专项债券资金情况",offset = 2,nullable = true)
    private String w14;
    // 3.支付进度
    @ExcelTitleBind(positionTitle = "来地方政府专项债券资金情况",offset = 3,nullable = true)
    private String w15;

    //是否申请前市级期费
    @ExcelTitleBind(title = "是否申请前市级期费",nullable = true)
    private String w16;

    /**前期费项目 协议还款时间*/
    @ExcelTitleBind(title = "协议还款时间")
    private Date cityAgreedRepaymentTime;

    /**前期费项目 是否归还前期费*/
    @ExcelTitleBind(title = "前期费项目是否归还前期费")
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    private String cityWhetherToReturnTheEarlyFee;

    /**前期费项目 拨付金额（万元）*/
    @ExcelTitleBind(title = "拨付金额（万元）")
    private String cityAmountAllocated;

    /**前期费项目 未还款金额（万元）*/
    @ExcelTitleBind(title = "未还款金额（万元）")
    private String cityOutstandingAmount;

    /** 项目业主 */
    @ExcelTitleBind(title = "项目业主")
    private String enterpriseName;

    /** 负责人及联系电话 */
    @ExcelTitleBind(title = "负责人及联系电话")
    private String contactNameAndPhone;

    /**标签*/
    @ExcelTitleBind(title = "标签")
    private String tags;

    /** 备注 */
    @ExcelTitleBind(title = "备注")
    private String remark;

}
