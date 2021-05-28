package top.yumbo.test.excel;


import lombok.Data;
import top.yumbo.excel.annotation.ExcelCellBindAnnotation;
import top.yumbo.excel.annotation.ExcelTableHeaderAnnotation;
import top.yumbo.excel.util.BigDecimalUtils;
import top.yumbo.excel.util.ExcelImportExportUtils;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/20 14:18
 */
@Data
@ExcelTableHeaderAnnotation(height = 4, tableName = "区域季度数据")// 表头占4行
public class RegionQuarterETLSyncResponse {


    /**
     * 年份
     */
    @ExcelCellBindAnnotation(title = "时间", exception = "年份格式不正确", pattern = "([0-9]{4})年")
    private Integer year;

    /**
     * 季度，填写1到4的数字
     */
    @ExcelCellBindAnnotation(title = "时间", exception = "季度只能是1，2，3，4", pattern = "([1-4]{1})季")
    private Integer quarter;

    /**
     * 地区代码，存储最末一级的地区代码就可以
     */
    @ExcelCellBindAnnotation(title = "地区", width = 2, exception = "地区不存在")
    private String regionCode;

    /**
     * 违约主体家数
     */
    @ExcelCellBindAnnotation(title = "违约主体家数", exception = "数值格式不正确")
    private Integer breachNumber;

    /**
     * 合计违约规模
     */
    @ExcelCellBindAnnotation(title = "合计违约规模", exception = ExcelImportExportUtils.INCORRECT_FORMAT_EXCEPTION, size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING)
    private BigDecimal breachTotalScale;

    /**
     * 风险性质 字典1260
     */
    @ExcelCellBindAnnotation(title = "风险性质", exception = ExcelImportExportUtils.NOT_BLANK_EXCEPTION)
    private String riskNature;

    /**
     * 风险品种 字典1261
     */
    @ExcelCellBindAnnotation(title = "风险品种", exception = ExcelImportExportUtils.NOT_BLANK_EXCEPTION)
    private String riskVarieties;

    /**
     * 区域偿债统筹管理能力 是否字典1022
     */
    @ExcelCellBindAnnotation(title = "区域偿债统筹管理能力", exception = ExcelImportExportUtils.NOT_BLANK_EXCEPTION)
    private String regionDebtManage;

    /**
     * 区域内私募可转债历史信用记录 是否字典1022
     */
    @ExcelCellBindAnnotation(title = "区域内私募可转债历史信用记录", exception = ExcelImportExportUtils.NOT_BLANK_EXCEPTION)
    private String calBondsHistoryCredit;

    /**
     * 还款可协调性 强弱字典1259
     */
    @ExcelCellBindAnnotation(title = "还款可协调性", exception = ExcelImportExportUtils.NOT_BLANK_EXCEPTION)
    private String repayCoordinated;

    /**
     * 业务合作可协调性 强弱字典1259
     */
    @ExcelCellBindAnnotation(title = "业务合作可协调性", exception = ExcelImportExportUtils.NOT_BLANK_EXCEPTION)
    private String cooperationCoordinated;

    /**
     * 数财通系统部署情况 是否字典1022
     */
    @ExcelCellBindAnnotation(title = "数财通系统部署情况", exception = ExcelImportExportUtils.NOT_BLANK_EXCEPTION)
    private String sctDeployStatus;


}
