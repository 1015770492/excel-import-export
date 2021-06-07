package top.yumbo.excel.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author jinhua
 * @date 2021/5/21 15:53
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCellBind {
    /**
     * 绑定的标题名称，
     * 通过扫描单元格表头可以确定表头所在的索引列，然后在根据width就能确定单元格
     */
    String title() default "";

    /**
     * 单元格宽度，对于合并单元格的处理
     * 确定表格的位置采用： 下标（解析过程会得到下标） + 单元格的宽度
     * 这样就可以确定单元格的位子和占据的宽度
     */
    int width() default 1;

    /**
     * 注入的异常消息，为了校验单元格内容
     * 校验失败应该返回的消息提升
     */
    String exception() default "";

    /**
     * 规模，对于BigDecimal类型的需要进行转换
     */
    String size() default "1";

    /**
     * 正则截取单元格内容
     * 一个单元格中的部分内容，例如 2020年2季度，只想单独取出年、季度这两个数字
     */
    String importPattern() default "";

    /**
     * 导出的字符串格式化填入，利用StringFormat.format进行字符串占位和替换
     */
    String exportFormat() default "";

    /**
     * 导出功能，该字段可能是多个单元格的内容（连续单元格），按照split拆分和填充。默认逗号
     */
    String exportSplit() default "";

    /**
     * 合并多个字段的顺序，多个字段构成一个标题，例如时间 年+季度
     */
    int exportPriority() default 0;

    /**
     * 默认不可以为空
     */
    boolean nullable() default false;

    /**
     * 单元格索引位置
     */
    int index() default -1;




}
