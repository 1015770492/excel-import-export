package top.yumbo.test.excel.importDemo;

//import io.swagger.annotations.ApiModelProperty;

import lombok.Data;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.annotation.ExcelTitleBind;
import top.yumbo.excel.annotation.MapEntry;

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
//    /** 行业分类 */
//    @ExcelTitleBind(title = "行业分类")
//    private String industry;
//
//    /** 委内责任处室 */
//    @ExcelTitleBind(title = "委内责任处室")
//    private String responsibleDepartment;
//
//
//
//    /** 开工以来累计完成投资 （万元） */
//    @ExcelTitleBind(title = "累计完成投资（万元）")
//    private String allCumulativeInvestmentAmount;
//
//    /** 本年累计完成投资 */
//    @ExcelTitleBind(title = "本年累计完成投资")
//    private String thisYearCumulativeInvestmentAmount;
//
//    /** 剩余投资 （万元） */
//    @ExcelTitleBind(title = "剩余投资（万元）")
//    private String leftInvestmentAmount;
//
//
//
//    /** 录入月份 */
//    @ExcelTitleBind(title = "录入月份")
//    private String inputMonth;
//
//    /** 本月完成投资 */
//    @ExcelTitleBind(title = "本月完成投资")
//    private String thisMonthFinishInvestmentAmount;
//
//    /** 是否为省重大项目 */
//    @ExcelTitleBind(title = "是否为省重大项目")
//    @MapEntry(key = "是", value = "1")
//    @MapEntry(key = "否", value = "0")
//    private String provinceKeyProject;
//
//    /** 是否为省重中之重项目 */
//    @ExcelTitleBind(title = "是否为省重中之重项目")
//    @MapEntry(key = "是", value = "1")
//    @MapEntry(key = "否", value = "0")
//    private String provinceMostKeyProject;
//
//    /** 是否为市重大项目 */
//    @ExcelTitleBind(title = "是否为市重大项目")
//    @MapEntry(key = "是", value = "1")
//    @MapEntry(key = "否", value = "0")
//    private String cityKeyProject;
//
//    /** 是否为市重中之重项目 */
//    @ExcelTitleBind(title = "是否为市重中之重项目")
//
//    @MapEntry(key = "是", value = "1")
//    @MapEntry(key = "否", value = "0")
//    private String cityMostKeyProject;
//
//    /**是否为现代化基础设施建设项目*/
//    @ExcelTitleBind(title = "是否为现代化基础设施建设项目")
//    @MapEntry(key = "是", value = "1")
//    @MapEntry(key = "否", value = "0")
//    private String isModernInfrastuctureProject;
//
//    /**是否为现代化基础设施建设项目*/
//    @ExcelTitleBind(title = "是否为现代化基础设施建设项目" )
//    private String modernInfrastuctureProject;
//
//    /** 是否申报集中开工 */
//    @ExcelTitleBind(title = "是否申报集中开工")
//    @MapEntry(key = "是", value = "1")
//    @MapEntry(key = "否", value = "0")
//    private String concentratedStart;
//
//    /** 集中开工年份批次 */
//    @ExcelTitleBind(title = "集中开工年份批次")
//    private String concentratedBatch;
//
//    /** 是否已开工 */
//    @ExcelTitleBind(title = "是否已开工")
//    @MapEntry(key = "是", value = "1")
//    @MapEntry(key = "否", value = "0")
//    private String constructionStarted;
//
//
//
//
//
//    /** 入库时间 */
//    @ExcelTitleBind(title = "入库时间")
//    private String regulatedTime;
//
//
//
//    /** 存在的问题和困难 */
//    @ExcelTitleBind(title = "存在的问题和困难")
//    private String problems;
//
//    /** 需协调解决事项 */
//    @ExcelTitleBind(title = "需协调解决事项")
//    private String coordinateResolutionProblem;
//
//    /** 协调解决事项层级 */
//    @ExcelTitleBind(title = "协调解决事项层级")
//    private String coordinateResolutionLevel;
//
//    /** 需协调解决部门 */
//    @ExcelTitleBind(title = "需协调解决部门")
//    private String coordinateResolutionDepartment;
//
//    /** 办理情况 */
//    @ExcelTitleBind(title = "办理情况")
//    private String handleResult;
//
//    /** 是否需信贷支持 */
//    @ExcelTitleBind(title = "是否需信贷支持")
//    @MapEntry(key = "是", value = "1")
//    @MapEntry(key = "否", value = "0")
//    private String needCreditSupport;
//
//
//
//
//
//
//
//    /** 是否申请市级前期费 */
//    @ExcelTitleBind(title = "是否申请市级前期费")
//    @MapEntry(key = "是", value = "1")
//    @MapEntry(key = "否", value = "0")
//    private String applyCityFund;
//
//    /** 地方政府专项债券资金情况下达年份批次 */
//    @ExcelTitleBind(title = "地方政府专项债券资金情况下达年份批次")
//    private String localGovBatch;
//    /** 地方政府专项债券资金情况资金类型 */
//    @ExcelTitleBind(title = "地方政府专项债券资金情况资金类型")
//    private String localGovFoundType;
//    /** 地方政府专项债券资金情况下达资金（万元） */
//    @ExcelTitleBind(title = "地方政府专项债券资金情况下达资金（万元）")
//    private String localGovTotalAmount;
//    /** 地方政府专项债券资金情况支付资金（万元） */
//    @ExcelTitleBind(title = "地方政府专项债券资金情况支付资金（万元）")
//    private String localGovPayAmount;
//    /** 地方政府专项债券资金情况支付进度 */
//    @ExcelTitleBind(title = "地方政府专项债券资金情况支付进度")
//    private String localGovProgress;
//    /** 中央预算内投资资金情况下达年份批次 */
//    @ExcelTitleBind(title = "中央预算内投资资金情况下达年份批次")
//    private String centerBatch;
//    /** 中央预算内投资资金情况资金类型 */
//    @ExcelTitleBind(title = "中央预算内投资资金情况资金类型")
//    private String centerFoundType;
//    /** 中央预算内投资资金情况下达资金（万元） */
//    @ExcelTitleBind(title = "中央预算内投资资金情况下达资金（万元）")
//    private String centerTotalAmount;
//    /** 中央预算内投资资金情况支付资金（万元） */
//    @ExcelTitleBind(title = "中央预算内投资资金情况支付资金（万元）")
//    private String centerPayAmount;
//    /** 中央预算内投资资金情况支付进度 */
//    @ExcelTitleBind(title = "中央预算内投资资金情况支付进度")
//    private String centerProgress;
//    /** 省预算内投资资金情况下达年份批次 */
//    @ExcelTitleBind(title = "省预算内投资资金情况下达年份批次")
//    private String provinceBatch;
//    /** 省预算内投资资金情况资金类型 */
//    @ExcelTitleBind(title = "省预算内投资资金情况资金类型")
//    private String provinceFoundType;
//    /** 省预算内投资资金情况下达资金（万元） */
//    @ExcelTitleBind(title = "省预算内投资资金情况下达资金（万元）")
//    private String provinceTotalAmount;
//    /** 省预算内投资资金情况支付资金（万元） */
//    @ExcelTitleBind(title = "省预算内投资资金情况支付资金（万元）")
//    private String provincePayAmount;
//    /** 省预算内投资资金情况支付进度 */
//    @ExcelTitleBind(title = "省预算内投资资金情况支付进度")
//    private String provinceProgress;
//
//
//    /** 前期费项目下达年份批次 */
//    @ExcelTitleBind(title = "前期费项目下达年份批次")
//    private String cityBatch;
//
//
//
//
//
//
//
//
//    /** 前期费项目资金类型 */
//    @ExcelTitleBind(title = "前期费项目资金类型")
//    private String cityFoundType;
//
//    /** 前期费项目下达资金（万元） */
//    @ExcelTitleBind(title = "前期费项目下达资金（万元）")
//    private String cityTotalAmount;
//
//    /** 前期费项目支付资金（万元） */
//    @ExcelTitleBind(title = "前期费项目支付资金（万元）")
//    private String cityPayAmount;
//
//    /** 前期费项目支付进度 */
//    @ExcelTitleBind(title = "前期费项目支付进度")
//    private String cityProgress;



}
