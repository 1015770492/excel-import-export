package top.yumbo.test.excel.importDemo;


import lombok.Data;
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelTableHeader;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/20 14:18
 */
@Data
@ExcelTableHeader(height = 4, tableName = "区域季度数据")// 表头占4行
public class ImportForQuarter {

    /**
     * 年份，为了避免暴露一些隐秘消息故字段都采用了w命名，防止泄露机密。不影响结果
     */
    // 根据正则截取单元格内容关于年份的值。其中exportFormat是导出excel填充到单元格的内容
    @ExcelCellBind(title = "时间", importPattern = "([0-9]{4})年", exportFormat = "$0年")
    private Integer w1;
    // 根据正则截取季度的数值，exportPriority是导出的顺序默认值是0，目的是与相同的title进行拼串，得到导出完整的单元格信息。
    // 在本次案例中目的是为了拼串成  $0年$1季，其中的$0被字段w1的值替换，$1被字段w2的值替换
    @ExcelCellBind(title = "时间", importPattern = "([1-4]{1})季", exportFormat = "第$1季", exportPriority = 1)
    private Integer w2;
    // 下面的exportSplit是导出功能需要用到的
    @ExcelCellBind(title = "地区", width = 2, exportSplit = ",", exportFormat = "$0,$1")
    private String w3;
    // 默认的异常消息就是格式不正确，如果在导入过程中出现不合法数据例如类型转换，单元格为空，会抛异常消息，提示你哪一行数据有问题
    @ExcelCellBind(title = "违约主体家数", exception = "格式不正确")
    private Integer w4;
    // 单位用size进行设置，例如表格上标注的单位是亿，这里的size就是下面的值。如果单位是%则填入字符串0.01即可以此类推
    @ExcelCellBind(title = "合计违约规模",size = "100000000")
    private BigDecimal w5;
    @ExcelCellBind(title = "风险性质", exception = "自定义的异常消息内容")
    private String w6;
    // nullable表示该字段是否为空，默认值是false。设置为true的情况下单元格内容如果为空这个字段的值就是null
    // 默认是不允许空的，故不设置为true的情况下，单元格内容为空则会抛异常并且提示第几行出错
    @ExcelCellBind(title = "风险品种",nullable = true)
    private String w7;
    @ExcelCellBind(title = "区域偿债统筹管理能力")
    private String w8;
    @ExcelCellBind(title = "区域内私募可转债历史信用记录")
    private String w9;
    @ExcelCellBind(title = "还款可协调性")
    private String w10;
    @ExcelCellBind(title = "业务合作可协调性")
    private String w12;
    @ExcelCellBind(title = "系统部署情况")
    private String w13;
}
