package top.yumbo.excel.test.importDemo;

import top.yumbo.excel.test.entity.ImportForQuarter;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.List;
import java.util.function.Consumer;

/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
public class ImportForQuarter_Consumer_Demo {
    /**
     * 将excel转换为List类型的数据 示例代码
     */
    public static void main(String[] args) throws Exception {


        System.out.println("=====导入季度数据======");
        final long start = System.currentTimeMillis();
//        String areaQuarter = "src/test/resources/excel/ImportForQuarter.xlsx";
        String areaQuarter = "src/test/resources/excel/ImportForQuarter_big.xlsx";
/*
        // 将内部分段处理的数据进行消费
        ExcelImportExportUtils.importExcelConsumer(new FileInputStream(areaQuarter), ImportForQuarter.class, (list) -> {
            if (Objects.requireNonNull(list).size() > 0) {
                list.forEach(System.out::println);
                System.out.println("总共有" + list.size() + "条记录");
            }
        }, 10000);
*/
        Consumer<List<ImportForQuarter>> consumer = (list) -> {};
        // 从consumer中得到List<ImportForQuarter> 或者 ImportForQuarter


//        test(consumer);

        final long end = System.currentTimeMillis();
        System.out.println("总共耗时" + (end - start) + "毫秒");

    }

    static <T> void test(Consumer<List<T>> consumer) {

    }
    public abstract class TypeReference<T> implements Comparable<TypeReference<T>>
    {
        protected final Type _type;

        protected TypeReference()
        {
            Type superClass = getClass().getGenericSuperclass();
            if (superClass instanceof Class<?>) { // sanity check, should never happen
                throw new IllegalArgumentException("Internal error: TypeReference constructed without actual type information");
            }

            _type = ((ParameterizedType) superClass).getActualTypeArguments()[0];
        }

        public Type getType() { return _type; }

        /**
         * The only reason we define this method (and require implementation
         * of <code>Comparable</code>) is to prevent constructing a
         * reference without type information.
         */
        @Override
        public int compareTo(TypeReference<T> o) { return 0; }
        // just need an implementation, not a good one... hence ^^^
    }

}
