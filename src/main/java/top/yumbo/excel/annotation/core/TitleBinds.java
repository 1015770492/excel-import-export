package top.yumbo.excel.annotation.core;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 *
 *
 * @author 诗水人间
 * @link 博客:{https://yumbo.blog.csdn.net/}
 * @link github:{https://github.com/1015770492}
 * @link 在线文档:{https://1015770492.github.io/excel-import-export/}
 * @date 2021/9/1 22:04
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface TitleBinds {
    TitleBind[] value();
}
