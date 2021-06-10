package top.yumbo.test.excel.exportDemo;

import lombok.Data;
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.enums.ExceptionMsg;

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
    private Integer w1;
    @ExcelCellBind(title = "时间", importPattern = "([1-4]{1})季", exportFormat = "第$1季", exportPriority = 1)
    private Integer w2;
    @ExcelCellBind(title = "地区", width = 2, exportSplit = ",", exportFormat = "$0,$1")
    private String w3;
    @ExcelCellBind(title = "违约主体家数", exception = "数值格式不正确")
    private Integer w4;
    @ExcelCellBind(title = "合计违约规模"/*,size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING*/)
    private BigDecimal w5;
    @ExcelCellBind(title = "风险性质", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String w6;
    @ExcelCellBind(title = "风险品种", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String w7;
    @ExcelCellBind(title = "区域偿债统筹管理能力", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String w8;
    @ExcelCellBind(title = "区域内私募可转债历史信用记录", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String w9;
    @ExcelCellBind(title = "还款可协调性", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String w10;
    @ExcelCellBind(title = "业务合作可协调性", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String w12;
    @ExcelCellBind(title = "系统部署情况", exception = ExceptionMsg.NOT_BLANK_EXCEPTION)
    private String w13;

}
