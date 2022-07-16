package top.yumbo.test.excel.importDemo;

//import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.annotation.MapEntry;

/**
 * 投资项目导入类
 * @author: Haiming Yu
 * @createDate:2022/7/14
 * @description:
 */
@Data
@ExcelTableHeader(height = 8, sheetName = "项目总表")
public class ImportForInveProj {

    /** 审批监管平台代码 */
    @ExcelCellBind(title = "审批监管平台代码",nullable = true)
//    //@ApiModelProperty(value = "审批监管平台代码")
    private java.lang.String supervisionCode;

    /** 项目名称 */
    @ExcelCellBind(title = "项目名称")
    //@ApiModelProperty(value = "项目名称")
    private java.lang.String projectName;

    /** 建设地点 */
    @ExcelCellBind(title = "建设地点")
    //@ApiModelProperty(value = "建设地点")
    private java.lang.String constructionAddress;

    /** 所属批次 */
    @ExcelCellBind(title = "所属批次")
    //@ApiModelProperty(value = "所属批次")
    private java.lang.String belongBatch;

    /** 行业分类 */
    @ExcelCellBind(title = "行业分类")
    //@ApiModelProperty(value = "行业分类")
    private java.lang.String industry;

    /** 建设内容及规模 */
    @ExcelCellBind(title = "建设内容及规模")
    //@ApiModelProperty(value = "建设内容及规模")
    private java.lang.String constructionContent;

    /** 项目推进情况 */
    @ExcelCellBind(title = "项目推进情况")
    //@ApiModelProperty(value = "项目推进情况")
    private java.lang.String progressDetail;

    /** 行业主管部门 */
    @ExcelCellBind(title = "行业主管部门")
    //@ApiModelProperty(value = "行业主管部门")
    private java.lang.String manageDepartment;

    /** 委内责任处室 */
    @ExcelCellBind(title = "委内责任处室")
    //@ApiModelProperty(value = "委内责任处室")
    private java.lang.String responsibleDepartment;

    /** 总投资 （万元） */
    @ExcelCellBind(title = "总投资（万元）")
    //@ApiModelProperty(value = "总投资（万元）")
    private java.lang.String totalInvestmentAmount;

    /** 开工以来累计完成投资 （万元） */
    @ExcelCellBind(title = "开工以来累计完成投资（万元）")
    //@ApiModelProperty(value = "开工以来累计完成投资（万元）")
    private java.lang.String allCumulativeInvestmentAmount;

    /** 本年累计完成投资 */
    @ExcelCellBind(title = "本年累计完成投资")
    //@ApiModelProperty(value = "本年累计完成投资")
    private java.lang.String thisYearCumulativeInvestmentAmount;

    /** 剩余投资 （万元） */
    @ExcelCellBind(title = "剩余投资（万元）")
    //@ApiModelProperty(value = "剩余投资（万元）")
    private java.lang.String leftInvestmentAmount;

    /** 2022年计划完成投资 （万元） */
    @ExcelCellBind(title = "2022年计划完成投资（万元）")
    //@ApiModelProperty(value = "2022年计划完成投资（万元）")
    private java.lang.String thisYearPlanInvestmentAmount;

    /** 录入月份 */
    @ExcelCellBind(title = "录入月份")
    //@ApiModelProperty(value = "录入月份")
    private java.lang.String inputMonth;

    /** 本月完成投资 */
    @ExcelCellBind(title = "本月完成投资")
    //@ApiModelProperty(value = "本月完成投资")
    private java.lang.String thisMonthFinishInvestmentAmount;

    /** 是否为省重大项目 */
    @ExcelCellBind(title = "是否为省重大项目")

    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否为省重大项目")
    private java.lang.String provinceKeyProject;

    /** 是否为省重中之重项目 */
    @ExcelCellBind(title = "是否为省重中之重项目")

    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否为省重中之重项目")
    private java.lang.String provinceMostKeyProject;

    /** 是否为市重大项目 */
    @ExcelCellBind(title = "是否为市重大项目")

    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否为市重大项目")
    private java.lang.String cityKeyProject;

    /** 是否为市重中之重项目 */
    @ExcelCellBind(title = "是否为市重中之重项目")

    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否为市重中之重项目")
    private java.lang.String cityMostKeyProject;

    /**是否为现代化基础设施建设项目*/
    @ExcelCellBind(title = "是否为现代化基础设施建设项目")
    //@ApiModelProperty(value = "是否为现代化基础设施建设项目")

    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    private java.lang.String isModernInfrastuctureProject;

    /**是否为现代化基础设施建设项目*/
    @ExcelCellBind(title = "是否为现代化基础设施建设项目" )
    //@ApiModelProperty(value = "是否为现代化基础设施建设项目")
    private java.lang.String modernInfrastuctureProject;

    /** 是否申报集中开工 */
    @ExcelCellBind(title = "是否申报集中开工")

    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否申报集中开工")
    private java.lang.String concentratedStart;

    /** 集中开工年份批次 */
    @ExcelCellBind(title = "集中开工年份批次")
    //@ApiModelProperty(value = "集中开工年份批次")
    private java.lang.String concentratedBatch;

    /** 是否已开工 */
    @ExcelCellBind(title = "是否已开工")

    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否已开工")
    private java.lang.String constructionStarted;

    /** 是否已入库 */
    @ExcelCellBind(title = "是否已入库",replaceAll = true)
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否已入库")
    private java.lang.String regulated;

    /** 开工时间 */
    @ExcelCellBind(title = "开工时间")
    //@ApiModelProperty(value = "开工时间")
    private java.lang.String constructionStartTime;

    /** 入库时间 */
    @ExcelCellBind(title = "入库时间")
    //@ApiModelProperty(value = "入库时间")
    private java.lang.String regulatedTime;

    /** 完工时间 */
    @ExcelCellBind(title = "完工时间")
    //@ApiModelProperty(value = "完工时间")
    private java.lang.String finishTime;

    /** 存在的问题和困难 */
    @ExcelCellBind(title = "存在的问题和困难")
    //@ApiModelProperty(value = "存在的问题和困难")
    private java.lang.String problems;

    /** 需协调解决事项 */
    @ExcelCellBind(title = "需协调解决事项")
    //@ApiModelProperty(value = "需协调解决事项")
    private java.lang.String coordinateResolutionProblem;

    /** 协调解决事项层级 */
    @ExcelCellBind(title = "协调解决事项层级")
    //@ApiModelProperty(value = "协调解决事项层级")
    private java.lang.String coordinateResolutionLevel;

    /** 需协调解决部门 */
    @ExcelCellBind(title = "需协调解决部门")
    //@ApiModelProperty(value = "需协调解决部门")
    private java.lang.String coordinateResolutionDepartment;

    /** 办理情况 */
    @ExcelCellBind(title = "办理情况")
    //@ApiModelProperty(value = "办理情况")
    private java.lang.String handleResult;

    /** 是否需信贷支持 */
    @ExcelCellBind(title = "是否需信贷支持")
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否需信贷支持")
    private java.lang.String needCreditSupport;

    /** 是否申请地方政府专项债券 */
    @ExcelCellBind(title = "是否申请地方政府专项债券")
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否申请地方政府专项债券")
    private java.lang.String applyLocalGovFund;

    /** 是否申请中央预算内投资 */
    @ExcelCellBind(title = "是否申请中央预算内投资")
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否申请中央预算内投资")
    private java.lang.String applyCentralFund;

    /** 是否申请省预算内投资 */
    @ExcelCellBind(title = "是否申请省预算内投资")
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否申请省预算内投资")
    private java.lang.String applyProvinceFund;

    /** 是否申请市级前期费 */
    @ExcelCellBind(title = "是否申请市级前期费")
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    //@ApiModelProperty(value = "是否申请市级前期费")
    private java.lang.String applyCityFund;

    /** 地方政府专项债券资金情况下达年份批次 */
    @ExcelCellBind(title = "地方政府专项债券资金情况下达年份批次")
    //@ApiModelProperty(value = "地方政府专项债券资金情况下达年份批次")
    private java.lang.String localGovBatch;
    /** 地方政府专项债券资金情况资金类型 */
    @ExcelCellBind(title = "地方政府专项债券资金情况资金类型")
    //@ApiModelProperty(value = "地方政府专项债券资金情况资金类型")
    private java.lang.String localGovFoundType;
    /** 地方政府专项债券资金情况下达资金（万元） */
    @ExcelCellBind(title = "地方政府专项债券资金情况下达资金（万元）")
    //@ApiModelProperty(value = "地方政府专项债券资金情况下达资金（万元）")
    private java.lang.String localGovTotalAmount;
    /** 地方政府专项债券资金情况支付资金（万元） */
    @ExcelCellBind(title = "地方政府专项债券资金情况支付资金（万元）")
    //@ApiModelProperty(value = "地方政府专项债券资金情况支付资金（万元）")
    private java.lang.String localGovPayAmount;
    /** 地方政府专项债券资金情况支付进度 */
    @ExcelCellBind(title = "地方政府专项债券资金情况支付进度")
    //@ApiModelProperty(value = "地方政府专项债券资金情况支付进度")
    private java.lang.String localGovProgress;
    /** 中央预算内投资资金情况下达年份批次 */
    @ExcelCellBind(title = "中央预算内投资资金情况下达年份批次")
    //@ApiModelProperty(value = "中央预算内投资资金情况下达年份批次")
    private java.lang.String centerBatch;
    /** 中央预算内投资资金情况资金类型 */
    @ExcelCellBind(title = "中央预算内投资资金情况资金类型")
    //@ApiModelProperty(value = "中央预算内投资资金情况资金类型")
    private java.lang.String centerFoundType;
    /** 中央预算内投资资金情况下达资金（万元） */
    @ExcelCellBind(title = "中央预算内投资资金情况下达资金（万元）")
    //@ApiModelProperty(value = "中央预算内投资资金情况下达资金（万元）")
    private java.lang.String centerTotalAmount;
    /** 中央预算内投资资金情况支付资金（万元） */
    @ExcelCellBind(title = "中央预算内投资资金情况支付资金（万元）")
    //@ApiModelProperty(value = "中央预算内投资资金情况支付资金（万元）")
    private java.lang.String centerPayAmount;
    /** 中央预算内投资资金情况支付进度 */
    @ExcelCellBind(title = "中央预算内投资资金情况支付进度")
    //@ApiModelProperty(value = "中央预算内投资资金情况支付进度")
    private java.lang.String centerProgress;
    /** 省预算内投资资金情况下达年份批次 */
    @ExcelCellBind(title = "省预算内投资资金情况下达年份批次")
    //@ApiModelProperty(value = "省预算内投资资金情况下达年份批次")
    private java.lang.String provinceBatch;
    /** 省预算内投资资金情况资金类型 */
    @ExcelCellBind(title = "省预算内投资资金情况资金类型")
    //@ApiModelProperty(value = "省预算内投资资金情况资金类型")
    private java.lang.String provinceFoundType;
    /** 省预算内投资资金情况下达资金（万元） */
    @ExcelCellBind(title = "省预算内投资资金情况下达资金（万元）")
    //@ApiModelProperty(value = "省预算内投资资金情况下达资金（万元）")
    private java.lang.String provinceTotalAmount;
    /** 省预算内投资资金情况支付资金（万元） */
    @ExcelCellBind(title = "省预算内投资资金情况支付资金（万元）")
    //@ApiModelProperty(value = "省预算内投资资金情况支付资金（万元）")
    private java.lang.String provincePayAmount;
    /** 省预算内投资资金情况支付进度 */
    @ExcelCellBind(title = "省预算内投资资金情况支付进度")
    //@ApiModelProperty(value = "省预算内投资资金情况支付进度")
    private java.lang.String provinceProgress;


    /** 前期费项目下达年份批次 */
    @ExcelCellBind(title = "前期费项目下达年份批次")
    //@ApiModelProperty(value = "前期费项目下达年份批次")
    private java.lang.String cityBatch;

    /**前期费项目协议还款时间*/
    @ExcelCellBind(title = "前期费项目协议还款时间")
    //@ApiModelProperty(value = "前期费项目协议还款时间")
    private java.util.Date cityAgreedRepaymentTime;

    /**前期费项目是否归还前期费*/
    @ExcelCellBind(title = "前期费项目是否归还前期费")
    @MapEntry(key = "是", value = "1")
    @MapEntry(key = "否", value = "0")
    private java.lang.String cityWhetherToReturnTheEarlyFee;

    /**前期费项目拨付金额（万元）*/
    @ExcelCellBind(title = "前期费项目拨付金额（万元）")
    //@ApiModelProperty(value = "前期费项目拨付金额（万元）")
    private java.lang.String cityAmountAllocated;

    /**前期费项目未还款金额（万元）*/
    @ExcelCellBind(title = "前期费项目未还款金额（万元）")
    //@ApiModelProperty(value = "前期费项目未还款金额（万元）")
    private java.lang.String cityOutstandingAmount;

    /** 前期费项目资金类型 */
    @ExcelCellBind(title = "前期费项目资金类型")
    //@ApiModelProperty(value = "前期费项目资金类型")
    private java.lang.String cityFoundType;

    /** 前期费项目下达资金（万元） */
    @ExcelCellBind(title = "前期费项目下达资金（万元）")
    //@ApiModelProperty(value = "前期费项目下达资金（万元）")
    private java.lang.String cityTotalAmount;

    /** 前期费项目支付资金（万元） */
    @ExcelCellBind(title = "前期费项目支付资金（万元）")
    //@ApiModelProperty(value = "前期费项目支付资金（万元）")
    private java.lang.String cityPayAmount;

    /** 前期费项目支付进度 */
    @ExcelCellBind(title = "前期费项目支付进度")
    //@ApiModelProperty(value = "前期费项目支付进度")
    private java.lang.String cityProgress;




    /** 项目业主 */
    @ExcelCellBind(title = "项目业主")
    //@ApiModelProperty(value = "项目业主")
    private java.lang.String enterpriseName;

    /** 负责人及联系电话 */
    @ExcelCellBind(title = "负责人及联系电话")
    //@ApiModelProperty(value = "负责人及联系电话")
    private java.lang.String contactNameAndPhone;

    /**标签*/
    @ExcelCellBind(title = "标签")
    //@ApiModelProperty(value = "标签")
    private java.lang.String tags;

    /** 备注 */
    @ExcelCellBind(title = "备注")
    //@ApiModelProperty(value = "备注")
    private java.lang.String remark;

}
