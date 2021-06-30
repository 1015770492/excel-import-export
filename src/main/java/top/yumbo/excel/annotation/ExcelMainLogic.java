package top.yumbo.excel.annotation;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author jinhua
 * @date 2021/5/21 15:53
 * excel主逻辑注解，用于标注主逻辑字段
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelMainLogic {
    /**
     * 主逻辑字段，默认收集所有字段，因为ExcelCellBind的logic默认值为0
     * 如果存在逻辑，例如 1,2,3 三个分支，选中一个后，对应的字段变成
     * 则会返回一个List，size大小为3
     */
    String value() default "0";
}
