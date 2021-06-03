package top.yumbo.test.excel.exportDemo.simpleDemo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.IndexedColors;
import top.yumbo.excel.interfaces.SerializableFunction;
import top.yumbo.excel.util.ReflectionUtil;

import java.lang.reflect.Field;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
public class ExcelSimpleExportDemo {
    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    static class User{
        private int age;
        private String name;
    }
    public static void main(String[] args) {
        final User user = new User(1,"z3");
//        final Supplier<String> getName = user::getName;
//        System.out.println(getName.get());
        //方法引用
        SerializableFunction<User, String> getName1 = User::getName;
        Field field = ReflectionUtil.getField(getName1);
        final IndexedColors red = IndexedColors.BLACK;
        System.out.println(red.index);
//        System.out.println(field.getName());
//        System.out.println(field.getType().getName());

    }
}
