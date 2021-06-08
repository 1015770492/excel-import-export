package top.yumbo.test.excel.exportDemo;

import lombok.Data;
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelCellStyle;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.enums.ExceptionMsg;
import top.yumbo.excel.util.BigDecimalUtils;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
@Data
@ExcelTableHeader(height = 4, tableName = "区域季度数据", resource = "path://src/test/java/top/yumbo/test/excel/2.xlsx")
// 表头占4行，使用了相对路径
public class ExportForQuarter {

    /**
     * 年份
     */
    @ExcelCellBind(title = "时间", importPattern = "([0-9]{4})年", exportFormat = "$0年")
    private Integer year;

    /**
     * 季度，填写1到4的数字
     */
    @ExcelCellBind(title = "时间", importPattern = "([1-4]{1})季", exportFormat = "第$1季", exportPriority = 1)
    private Integer quarter;

    /**
     * 地区代码，存储最末一级的地区代码就可以
     */
    @ExcelCellBind(title = "地区", width = 2, exportSplit = ",", exportFormat = "$0,$1")
    private String regionCode;

    /**
     * 违约主体家数
     */
    @ExcelCellBind(title = "违约主体家数", exception = "数值格式不正确")
    private Integer breachNumber;

    /**
     * 合计违约规模
     */
    @ExcelCellBind(title = "合计违约规模", size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING, exception = ExceptionMsg.INCORRECT_FORMAT_EXCEPTION)
    private BigDecimal breachTotalScale;

    /**
     * 风险性质 字典1260
     */
    @ExcelCellBind(title = "风险性质", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String riskNature;

    /**
     * 风险品种 字典1261
     */
    @ExcelCellBind(title = "风险品种", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String riskVarieties;

    /**
     * 区域偿债统筹管理能力 是否字典1022
     */
    @ExcelCellBind(title = "区域偿债统筹管理能力", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String regionDebtManage;

    /**
     * 区域内私募可转债历史信用记录 是否字典1022
     */
    @ExcelCellBind(title = "区域内私募可转债历史信用记录", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String calBondsHistoryCredit;

    /**
     * 还款可协调性 强弱字典1259
     */
    @ExcelCellBind(title = "还款可协调性", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String repayCoordinated;

    /**
     * 业务合作可协调性 强弱字典1259
     */
    @ExcelCellBind(title = "业务合作可协调性", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String cooperationCoordinated;

    /**
     * 数财通系统部署情况 是否字典1022
     */
    @ExcelCellBind(title = "数财通系统部署情况", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String sctDeployStatus;


}
