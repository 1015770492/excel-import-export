package top.yumbo.excel.annotation.core;


import java.lang.annotation.*;

/**
 * 绑定标题头注解
 *
 * @author 诗水人间
 * @link 博客:{https://yumbo.blog.csdn.net/}
 * @link github:{https://github.com/1015770492}
 * @link 在线文档:{https://1015770492.github.io/excel-import-export/}
 * @date 2021/9/1 22:04
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(ExcelTitleBinds.class)
public @interface ExcelTitleBind {
    /**
     * 绑定的标题名称，
     * 通过扫描单元格表头可以确定表头所在的索引列，然后在根据width就能确定单元格
     */
    String title() default "";

    /**
     * 合并连续的单元格，例如一个标题占据两列的情况
     */
    int width() default 1;

    /**
     * 单位基数，对于BigDecimal类型的需要进行转换
     */
    String size() default "1";

    /**
     * 异常提醒
     */
    String exception() default "格式不正确";

    /**
     * 正则导入
     * 一个单元格中的部分内容，例如 2020年2季度，只想单独取出年（2020）、季度（2）这两个数字
     */
    String importPattern() default "";

    /**
     * 自定义正则分隔符
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
    boolean nullable() default true;

    /**
     * 单元格索引位置
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
