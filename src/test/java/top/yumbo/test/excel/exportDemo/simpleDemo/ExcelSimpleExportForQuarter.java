package top.yumbo.test.excel.exportDemo.simpleDemo;

import lombok.Data;
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.enumeration.ExceptionMsg;
import top.yumbo.excel.util.BigDecimalUtils;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
@Data
@ExcelTableHeader(height = 4, tableName = "区域季度数据")// 表头占4行
public class ExcelSimpleExportForQuarter {

    /**
     * 年份
     */
    @ExcelCellBind(title = "时间", exportFormat = "%年", exportPriority = 1)
    private Integer year;

    /**
     * 季度，填写1到4的数字
     */
    @ExcelCellBind(title = "时间", exportFormat = "第%s季", exportPriority = 2)
    private Integer quarter;

    /**
     * 地区代码，存储最末一级的地区代码就可以
     */
    @ExcelCellBind(title = "地区", width = 2, exportSplit = ",")
    private String regionCode;

    /**
     * 违约主体家数
     */
    @ExcelCellBind(title = "违约主体家数", exception = "数值格式不正确")
    private Integer breachNumber;

    /**
     * 合计违约规模
     */
    @ExcelCellBind(title = "合计违约规模", exception = ExceptionMsg.INCORRECT_FORMAT_EXCEPTION, size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
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
