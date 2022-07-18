package top.yumbo.test.excel.exportDemo;

import lombok.Data;
import org.hibernate.validator.constraints.NotEmpty;
import top.yumbo.excel.annotation.ExcelTitleBind;
import top.yumbo.excel.annotation.ExcelTableHeader;

import javax.validation.constraints.Max;
import javax.validation.constraints.Min;
import java.math.BigDecimal;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
@Data
// 表头占4行，height，同时使用resource设置模板文件位置
@ExcelTableHeader(height = 4, sheetName = "区域季度数据", resource = "path:///D:/IdeaProjects/excel-import-export/src/test/java/top/yumbo/test/excel/2.xlsx")
// 如果关于path的设置，建议使用绝对路径，例如：上面的案例 以/开头，如果要配置相对路径，需要确保模板excel文件与.class文件的路径关系
// 本地项目中容易导致一个错误就算编译后的.class文件中会发现模板文件不在编译的路径，就会导致路径错误
public class ExportForQuarter {

    /**
     * 年份，为了避免暴露一些隐秘消息故字段都采用了w命名，防止泄露机密。不影响结果
     */
    // 根据正则截取单元格内容关于年份的值。其中exportFormat是导出excel填充到单元格的内容
    @ExcelTitleBind(title = "时间", exportFormat = "$0年")
    @Min(value = 2017,message = "最小年份是2017年")
    @Max(value = 2021,message = "最大年份是2021年")
    private Integer w1;
    // 根据正则截取季度的数值，exportPriority是导出的顺序默认值是0，目的是与相同的title进行拼串，得到导出完整的单元格信息。
    // 在本次案例中目的是为了拼串成  $0年$1季，其中的$0被字段w1的值替换，$1被字段w2的值替换
    @ExcelTitleBind(title = "时间", exportFormat = "第$1季", exportPriority = 1)
    private Integer w2;
    // 下面的exportSplit是导出功能需要用到的
    @ExcelTitleBind(title = "地区", width = 2, exportSplit = ",", exportFormat = "$0,$1")
    @NotEmpty(message = "季度信息不能为空")
    private String w3;
    // 默认的异常消息就是格式不正确，如果在导入过程中出现不合法数据例如类型转换，单元格为空，会抛异常消息，提示你哪一行数据有问题
    @ExcelTitleBind(title = "违约主体家数", exception = "格式不正确")
    private Integer w4;
    // 单位用size进行设置，例如表格上标注的单位是亿，这里的size就是下面的值。如果单位是%则填入字符串0.01即可以此类推
    @ExcelTitleBind(title = "合计违约规模",size = "100000000")
    private BigDecimal w5;
    @ExcelTitleBind(title = "风险性质", exception = "自定义的异常消息内容")
    private String w6;
    // nullable表示该字段是否为空，默认值是false。设置为true的情况下单元格内容如果为空这个字段的值就是null
    // 默认是不允许空的，故不设置为true的情况下，单元格内容为空则会抛异常并且提示第几行出错
    @ExcelTitleBind(title = "风险品种",nullable = true)
    private String w7;
    @ExcelTitleBind(title = "区域偿债统筹管理能力")
    private String w8;
    @ExcelTitleBind(title = "还款可协调性")
    private String w9;
    @ExcelTitleBind(title = "业务合作可协调性")
    private String w10;
    @ExcelTitleBind(title = "数财通系统部署情况")
    private String w11;

}
