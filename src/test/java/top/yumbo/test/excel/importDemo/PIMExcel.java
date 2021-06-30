package top.yumbo.test.excel.importDemo;

import lombok.Data;
import top.yumbo.excel.annotation.*;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/6/30 13:52
 * 不动资产导入excel请求
 */
@Data
@ExcelTableHeader(height = 1)
public class PIMExcel {

    @ExcelCellBind(title = "*抵押人/所有权人")
    private String mortgagorIm;

    @ExcelMainLogic
    @MapEntry(key = "土地使用权", value = "8001301")
    @MapEntry(key = "不动产", value = "8001302")
    @MapEntry(key = "在建工程", value = "8001303")
    @ExcelCellBind(title = "*抵押标的类型")
    private String mortgageSubjectTypeIm;

    @ExcelCellBind(title = "*权证证号")
    private String warrantNumberIm;

    @ExcelCellBind(title = "*坐落")
    private String siteIm;

    @MapEntry(key = "集体土地使用权", value = "8001601")
    @MapEntry(key = "房屋等建筑物、构筑物所有权", value = "8001602")
    @MapEntry(key = "森林、林木所有权", value = "8001603")
    @MapEntry(key = "耕地、林地、草地等土地承包经营权", value = "8001604")
    @MapEntry(key = "建设用地使用权", value = "8001605")
    @MapEntry(key = "宅基地使用权", value = "8001606")
    @MapEntry(key = "海域使用权", value = "8001607")
    @MapEntry(key = "地役权", value = "8001608")
    @MapEntry(key = "抵押权", value = "8001609")
    @MapEntry(key = "法律规定需要登记的其他不动产权利", value = "8001610")
    @ExcelCellBind(title = "*权利类型")
    private String rightTypeIm;

    @ExcelCellBind(title = "*使用年限")
    private String usefulLifeIm;

    @ExcelCellBind(title = "*用途")
    private String purposeIm;

    @MapEntry(key = "否", value = "0")
    @MapEntry(key = "是", value = "1")
    @ExcelCellBind(title = "*是否受限")
    private String isLimitedIm;

    @MapEntry(key = "否", value = "0")
    @MapEntry(key = "是", value = "1")
    @ExcelCellBind(title = "*是否国营主要物业或标志性资产")
    private String isgywybzz;

    @ExcelCellBind(title = "*房屋建筑面积", nullable = true)
    private String buildingAreaIm;

    @ExcelCellBind(title = "*房屋建筑面积单位", nullable = true)
    private String buildingAreaImUnit;

    @ExcelFollowLogic(value = "土地使用权")
    @ExcelFollowLogic(value = "在建工程")
    @ExcelCellBind(title = "*宗地面积", nullable = true)
    private BigDecimal patriarchalAreaIm;

    @ExcelCellBind(title = "*宗地面积单位", nullable = true)
    private String patriarchalAreaImSize;

    @ExcelCellBind(title = "*抵押类型")
    private String patriarchalAreaImUnit;

    @ExcelCellBind(title = "*是否有评估价值")
    @MapEntry(key = "否", value = "0")
    @MapEntry(key = "是", value = "1")
    private String isAssessmenttitleIm;

    @ExcelCellBind(title = "*评估报告类型", nullable = true)
    private String assessmentReportTypeIm;

    @ExcelCellBind(title = "*评估机构", nullable = true)
    private String assessmentMechanismIm;

    @ExcelCellBind(title = "*评估报告名称", nullable = true)
    private String assessmentReportNameIm;

    @ExcelCellBind(title = "*评估报告编号", nullable = true)
    private String assessmentReportSnoIm;

    @ExcelCellBind(title = "*评估价值", nullable = true)
    private BigDecimal assessmenttitleIm;
    @ExcelCellBind(title = "*评估价值单位", nullable = true)
    private String assessmenttitleImSize;

    @ExcelCellBind(title = "*评估基准日", nullable = true)
    private String assessmentBaseDateIm;

    @ExcelCellBind(title = "*权利性质")
    private String mortgageTypeIm;

    @ExcelCellBind(title = "*抵押权人")
    private String mortgageHolderType;

    @ExcelCellBind(title = "*受权抵押权公司名称", nullable = true)
    private String mortgageEnterprise;

    @ExcelCellBind(title = "*受权抵押权公司法定代表/责任人", nullable = true)
    private String mortgageEnterpriseInstrepr;

    @ExcelCellBind(title = "*受权抵押权公司联系电话", nullable = true)
    private String mortgageEnterpriseTel;

    @ExcelCellBind(title = "*担保范围")
    private String guaranteeRange;

    @ExcelCellBind(title = "*特殊说明", nullable = true)
    private String guaranteeRangeRemark;

}
