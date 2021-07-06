package top.yumbo.excel.annotation;


import java.lang.annotation.*;

/**
 * @author jinhua
 * @date 2021/5/21 15:53
 * excel主逻辑注解，用于标注主逻辑字段
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface CheckNullLogic {

    /**
     * 字段名称
     */
    String follow() default "";

    /**
     * follow字段的值为value时必填
     */
    String[] values() default {};

    /**
     * 标题
     */
    String followTitle() default "";

    /**
     * 当前标题
     */
    String title() default "";

    /**
     * 消息
     */
    String message() default "";
}
