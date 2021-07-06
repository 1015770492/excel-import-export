package top.yumbo.test.excel.importDemo;

import lombok.Data;
import org.hibernate.validator.constraints.NotEmpty;
import top.yumbo.excel.annotation.*;

import java.math.BigDecimal;
import java.time.LocalDate;

/**
 * @author jinhua
 * @date 2021/6/30 13:52
 * 不动资产导入excel请求
 */
@Data
@ExcelTableHeader(height = 2)
public class PIMExcel {

    @ExcelCellBind(title = "*抵押人/所有权人")
    private String mortgagorIm;

    @MapEntry(key = "土地使用权", value = "8001301")
    @MapEntry(key = "不动产", value = "8001302")
    @MapEntry(key = "在建工程", value = "8001303")
    @ExcelCellBind(title = "*抵押标的类型")
    private String mortgageSubjectTypeIm;

    @ExcelCellBind(title = "*权证证号")
    private String warrantNumberIm;

    @ExcelCellBind(title = "*坐落")
    private String siteIm;

    @MapEntry(key = "1", value = "8001601")
    @MapEntry(key = "2", value = "8001602")
    @MapEntry(key = "3", value = "8001603")
    @MapEntry(key = "4", value = "8001604")
    @MapEntry(key = "5", value = "8001605")
    @MapEntry(key = "6", value = "8001606")
    @MapEntry(key = "7", value = "8001607")
    @MapEntry(key = "8", value = "8001608")
    @MapEntry(key = "9", value = "8001609")
    @MapEntry(key = "10", value = "8001610")
    @ExcelCellBind(title = "*权利类型", splitRegex = ",")
    private String rightTypeIm;

    @ExcelCellBind(title = "*使用年限")
    private String usefulLifeIm;

    @ExcelCellBind(title = "*用途")
    private String purposeIm;

    @ExcelCellBind(title = "*是否受限")
    private String isLimitedIm;


    @ExcelCellBind(title = "*是否国营主要物业或标志性资产")
    private String isgywybzz;

    @ExcelCellBind(title = "*房屋建筑面积", nullable = true)
    @CheckNullLogic(follow = "mortgageSubjectTypeIm", values = {"8001302", "8001303"})
    private BigDecimal buildingAreaIm;

    @MapEntry(key = "公顷", value = "10000")
    @MapEntry(key = "平方公里", value = "1000000")
    @MapEntry(key = "平方米", value = "1")
    @MapEntry(key = "亩", value = "666.66667")
    @ExcelCellBind(title = "*房屋建筑面积单位", nullable = true)
    @CheckNullLogic(follow = "mortgageSubjectTypeIm", values = {"8001302", "8001303"})
    @AccountBigDecimalValue(follow = "buildingAreaIm", decimalFormat = "#.##")
    private String buildingAreaImSize;

    @ExcelCellBind(title = "*房屋建筑面积单位", nullable = true)
    @MapEntry(key = "公顷", value = "Hectares")
    @MapEntry(key = "平方公里", value = "kilometers")
    @MapEntry(key = "平方米", value = "meters")
    @MapEntry(key = "亩", value = "acres")
    @CheckNullLogic(follow = "mortgageSubjectTypeIm", values = {"8001302", "8001303"})
    private String buildingAreaImUnit;

    @ExcelCellBind(title = "*宗地面积", nullable = true)
    @CheckNullLogic(follow = "mortgageSubjectTypeIm", values = "8001301")
    private BigDecimal patriarchalAreaIm;

    // 换算成平方米,冗余字段用于单位换算，不存数据库
    @MapEntry(key = "公顷", value = "10000")
    @MapEntry(key = "平方公里", value = "1000000")
    @MapEntry(key = "平方米", value = "1")
    @MapEntry(key = "亩", value = "666.66667")
    @ExcelCellBind(title = "*宗地面积单位", nullable = true)
    @CheckNullLogic(follow = "mortgageSubjectTypeIm", values = "8001301")
    @AccountBigDecimalValue(follow = "patriarchalAreaIm", decimalFormat = "#.##")
    private String patriarchalAreaImSize;

    @MapEntry(key = "公顷", value = "Hectares")
    @MapEntry(key = "平方公里", value = "kilometers")
    @MapEntry(key = "平方米", value = "meters")
    @MapEntry(key = "亩", value = "acres")
    @ExcelCellBind(title = "*宗地面积单位", nullable = true)
    @CheckNullLogic(follow = "mortgageSubjectTypeIm", values = "8001301")
    private String patriarchalAreaImUnit;

    @MapEntry(key = "共同抵押", value = "0")
    @MapEntry(key = "承诺抵押", value = "1")
    @ExcelCellBind(title = "*抵押类型")
    private String mortgageTypeIm;

    @ExcelCellBind(title = "*是否有评估价值")
    @NotEmpty(message = "是否有评估价值不允许为空")
    private String isAssessmentValueIm;


    @ExcelCellBind(title = "*评估报告类型", nullable = true)
    @CheckNullLogic(follow = "isAssessmentValueIm", values = "是")
    @MapEntry(key = "预评估报告",value = "8012201")
    @MapEntry(key = "正式评估报告",value = "8012202")
    private String assessmentReportTypeIm;

    @ExcelCellBind(title = "*评估机构", nullable = true)
    @CheckNullLogic(follow = "isAssessmentValueIm", values = "是")
    private String assessmentMechanismIm;

    @ExcelCellBind(title = "*评估报告名称", nullable = true)
    @CheckNullLogic(follow = "isAssessmentValueIm", values = "是")
    private String assessmentReportNameIm;

    @ExcelCellBind(title = "*评估报告编号", nullable = true)
    @CheckNullLogic(follow = "isAssessmentValueIm", values = "是")
    private String assessmentReportSnoIm;

    @CheckNullLogic(follow = "isAssessmentValueIm", values = "是")
    @ExcelCellBind(title = "*评估价值", nullable = true)
    private BigDecimal assessmentValueIm;
    // 金额单位换算
    @MapEntry(key = "元", value = "1")
    @MapEntry(key = "万", value = "10000")
    @MapEntry(key = "百万", value = "1000000")
    @MapEntry(key = "千万", value = "10000000")
    @MapEntry(key = "亿", value = "100000000")
    @ExcelCellBind(title = "*评估价值单位", nullable = true)
    @AccountBigDecimalValue(follow = "assessmentValueIm", decimalFormat = "#.##")
    private String assessmentValueImSize;

    @ExcelCellBind(title = "*评估价值单位", nullable = true)
    @MapEntry(key = "元", value = "One")
    @MapEntry(key = "万", value = "TenThoursand")
    @MapEntry(key = "百万", value = "Million")
    @MapEntry(key = "千万", value = "TenMillion")
    @MapEntry(key = "亿", value = "HundredMillion")
    @CheckNullLogic(follow = "isAssessmentValueIm", values = "是")
    private String assessmentValueImUnit;

    @ExcelCellBind(title = "*评估基准日", nullable = true)
    @CheckNullLogic(follow = "isAssessmentValueIm", values = "是")
    private LocalDate assessmentBaseDateIm;

    @MapEntry(key = "出让", value = "0")
    @MapEntry(key = "划拨", value = "1")
    @ExcelCellBind(title = "*权利性质")
    private String rightProp;

    @MapEntry(key = "受托管理人",value = "0")
    @MapEntry(key = "受权抵押权人",value = "1")
    @ExcelCellBind(title = "*抵押权人")
    private String mortgageHolderType;

    @ExcelCellBind(title = "*受权抵押权公司名称", nullable = true)
    private String mortgageEnterprise;

    @ExcelCellBind(title = "*受权抵押权公司法定代表/责任人", nullable = true)
    private String mortgageEnterpriseInstrepr;

    @ExcelCellBind(title = "*受权抵押权公司联系电话", nullable = true)
    private String mortgageEnterpriseTel;

    @ExcelCellBind(title = "*担保范围")
    @MapEntry(key = "正常", value = "0")
    @MapEntry(key = "特殊", value = "1")
    private String guaranteeRange;

    @ExcelCellBind(title = "*特殊说明", nullable = true)
    private String guaranteeRangeRemark;
}
