package top.yumbo.test.excel.importDemo;


import lombok.Data;
import top.yumbo.excel.annotation.core.TitleBind;
import top.yumbo.excel.annotation.core.TableHeader;

import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/20 14:18
 */
@Data
@TableHeader(height = 4, tableName = "区域季度数据")
public class ImportForQuarter {

    /**
     * 年份，为了避免暴露一些隐秘消息故字段都采用了w命名，防止泄露机密。不影响结果
     */
    // 根据正则截取单元格内容关于年份的值。其中exportFormat是导出excel填充到单元格的内容
    @TitleBind(title = "时间", importPattern = "([0-9]{4})年")
    private Integer w1;
    @TitleBind(title = "时间", importPattern = "([1-4]{1})季")
    private Integer w2;
    // 下面的exportSplit是导出功能需要用到的
    @TitleBind(title = "地区", width = 2)
    private String w3;
    // 默认的异常消息就是格式不正确，如果在导入过程中出现不合法数据例如类型转换，单元格为空，会抛异常消息，提示你哪一行数据有问题
    @TitleBind(title = "违约主体家数", exception = "格式不正确")
    private Integer w4;
    // 单位用size进行设置，例如表格上标注的单位是亿，这里的size就是下面的值。如果单位是%则填入字符串0.01即可以此类推
    @TitleBind(title = "合计违约规模",size = "100000000")
    private BigDecimal w5;
    @TitleBind(title = "风险性质", exception = "自定义的异常消息内容")
    private String w6;
    // nullable表示该字段是否为空，默认值是false。设置为true的情况下单元格内容如果为空这个字段的值就是null
    // 默认是不允许空的，故不设置为true的情况下，单元格内容为空则会抛异常并且提示第几行出错
    @TitleBind(title = "风险品种",nullable = true)
    private String w7;
    @TitleBind(title = "区域偿债统筹管理能力")
    private String w8;
    @TitleBind(title = "还款可协调性")
    private String w9;
    @TitleBind(title = "业务合作可协调性")
    private String w10;
    @TitleBind(title = "数财通系统部署情况")
    private String w11;
}
