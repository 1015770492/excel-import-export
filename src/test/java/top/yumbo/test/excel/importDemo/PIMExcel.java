package top.yumbo.test.excel.importDemo;

import lombok.Data;
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.annotation.MapEntry;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/6/30 13:52
 * 不动资产导入excel请求
 */
@Data
@ExcelTableHeader(height = 3)
public class PIMExcel {

    @ExcelCellBind(title = "*抵押人/所有权人")
    private String mortgagorIm;

    @MapEntry(key = "")
    @ExcelCellBind(title = "*抵押标的类型")
    private String mortgageSubjectTypeIm;

    @ExcelCellBind(title = "*权证证号")
    private String warrantNumberIm;

    @ExcelCellBind(title = "*坐落")
    private String siteIm;

    @ExcelCellBind(title = "*权利类型")
    private String rightTypeIm;

    @ExcelCellBind(title = "*使用年限")
    private String usefulLifeIm;

    @ExcelCellBind(title = "*用途")
    private String purposeIm;

    @ExcelCellBind(title = "*是否受限")
    private String isLimitedIm;

    @ExcelCellBind(title = "*是否国营主要物业或标志性资产")
    private String isgywybzz;

    @ExcelCellBind(title = "*房屋建筑面积")
    private String buildingAreaIm;

    @ExcelCellBind(title = "*房屋建筑面积单位")
    private String buildingAreaImUnit;

    @ExcelCellBind(title = "*宗地面积")
    private BigDecimal patriarchalAreaIm;

    @ExcelCellBind(title = "*宗地面积单位")
    private String patriarchalAreaImSize;

    @ExcelCellBind(title = "*抵押类型")
    private String patriarchalAreaImUnit;

    @ExcelCellBind(title = "*是否有评估价值")
    private String isAssessmenttitleIm;

    @ExcelCellBind(title = "*评估报告类型")
    private String assessmentReportTypeIm;

    @ExcelCellBind(title = "*评估机构")
    private String assessmentMechanismIm;

    @ExcelCellBind(title = "*评估报告名称")
    private String assessmentReportNameIm;

    @ExcelCellBind(title = "*评估报告编号")
    private String assessmentReportSnoIm;

    @ExcelCellBind(title = "*评估价值")
    private BigDecimal assessmenttitleIm;
    @ExcelCellBind(title = "*评估价值单位")
    private String assessmenttitleImSize;

    @ExcelCellBind(title = "*评估基准日")
    private String assessmentBaseDateIm;

    @ExcelCellBind(title = "*权利性质")
    private String mortgageTypeIm;

    @ExcelCellBind(title = "*抵押权人")
    private String mortgageHolderType;

    @ExcelCellBind(title = "*受权抵押权公司名称")
    private String mortgageEnterprise;

    @ExcelCellBind(title = "*受权抵押权公司法定代表/责任人")
    private String mortgageEnterpriseInstrepr;

    @ExcelCellBind(title = "*受权抵押权公司联系电话")
    private String mortgageEnterpriseTel;

    @ExcelCellBind(title = "*担保范围")
    private String guaranteeRange;

    @ExcelCellBind(title = "*特殊说明")
    private String guaranteeRangeRemark;

}
