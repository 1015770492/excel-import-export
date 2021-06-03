package top.yumbo.excel.interfaces;

import top.yumbo.excel.util.ReflectionUtil;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;
import java.util.function.Predicate;

public class LambdaEnableCellStyle<T, R> {


    Map<SerializableFunction<T, R>, Predicate<T>> cache = new HashMap<>();

    public LambdaEnableCellStyle(SerializableFunction<T, R> serializableFunction, Predicate<T> predicate) {
        cache.put(serializableFunction, predicate);
    }

    LambdaEnableCellStyle and(SerializableFunction<T, R> serializableFunction, Predicate<T> predicate) {

        return this;
    }

    LambdaEnableCellStyle or(SerializableFunction<T, R> serializableFunction, Predicate predicate) {

        return this;
    }

    LambdaEnableCellStyle allOf(LambdaEnableCellStyle... lambdaEnableCellStyles) {

        return this;
    }

    LambdaEnableCellStyle anyOf(LambdaEnableCellStyle... lambdaEnableCellStyles) {

        return this;
    }

    boolean enableStyle(SerializableFunction<T, R> serializableFunction, T t, Predicate<T> predicate) {
        Field field = ReflectionUtil.getField(serializableFunction);
        field.getName();// 得到字段名称
        field.getType().getName();// 得到字段类型
        return predicate.test(t);
    }
}