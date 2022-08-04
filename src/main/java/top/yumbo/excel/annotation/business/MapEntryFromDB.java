package top.yumbo.excel.annotation.business;

import java.lang.annotation.*;

/**
 * @author jinhua
 * @date 2021/6/8 15:27
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface MapEntryFromDB {

    /**
     * sqlQuery或者MongoDB等数据库持久层应用
     * 得到字典，在导入的过程中进行字典映射。
     */
    String dbQuery() default "";

}
