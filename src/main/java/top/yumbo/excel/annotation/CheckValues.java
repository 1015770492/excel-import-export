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
public @interface CheckValues {

    /**
     * follow字段的值为value时必填
     */
    String[] values() default {};

    /**
     * 消息
     */
    String message() default "";
}
