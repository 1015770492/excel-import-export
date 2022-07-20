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
public @interface ExcelTitleBind {
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
    String exception() default "格式不正确";

    /**
     * 规模，对于BigDecimal类型的需要进行转换
     */
    String size() default "1";

    /**
     * 正则截取单元格部分内容，只需要部分其它内容丢掉
     * 一个单元格中的部分内容，例如 2020年2季度，只想单独取出年、季度这两个数字
     */
    String importPattern() default "";


    /**
     * 正则截取单元格内容，保留单元格内容，后面进行替换字典
     * 服务于replaceAllOrPart，如果使用了splitRegex，则会将内容切割进行replaceAllOrPart
     * 然后将将处理后的结果返回，然后再进行importPattern
     */
    String splitRegex() default "";

    /**
     * 包含字典key就完全替换为value
     * 例如：key=江西上饶, value=jx
     * replaceAll=true，那么就会被替换为jx。
     * 如果设置为false，只会替换字典部分内容，也就是变成：jx上饶
     */
    boolean replaceAll() default true;

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
     * 单元格索引位置，如果标题重复，没法通过标题来得到index，则可以通过 positionTitle() + offset() 来更新
     */
    String index() default "-1";

    /**
     * 重复标题的处理方案：利用不重复标题的位置 positionTitle + 当前标题到不重复标题的偏移offset量
     * 这样就可以不写死index
     * 定位
     */
    String positionTitle() default "";

    /**
     * 偏移
     */
    int offset() default 1;


}
