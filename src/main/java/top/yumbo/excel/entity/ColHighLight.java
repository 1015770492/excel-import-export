package top.yumbo.excel.entity;

import lombok.Builder;
import lombok.Data;
import top.yumbo.excel.interfaces.SerializableFunction;

import java.util.function.Predicate;

/**
 * @author jinhua
 * @date 2021/6/4 5:06
 */
@Data
@Builder
public class ColHighLight <T,R>{
    private SerializableFunction<T,R> function;

    private Predicate<T> predicate;
}
