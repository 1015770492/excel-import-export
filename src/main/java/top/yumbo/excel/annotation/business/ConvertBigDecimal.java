package top.yumbo.excel.annotation.business;

import java.lang.annotation.*;

/**
 * @author jinhua
 * @date 2021/7/4 13:09
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ConvertBigDecimal {
    /**
     * 给哪一个字段进行计算
     */
    String follow() default "";

    /**
     * bigDecimal格式化
     */
    String decimalFormat() default "";

}
