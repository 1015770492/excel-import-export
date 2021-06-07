package top.yumbo.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Map;
import java.util.function.Function;

/**
 * @author jinhua
 * @date 2021/6/4 11:08
 *
 * 标题-样式-断言器返回样式下标
 * 通过标题知道那个单元格，然后通过断言器判断是否需要启用样式
 */
@Data
@AllArgsConstructor
public class TitleCellStyleFunction<T,R> {

    /**
     * 哪一个标题的样式
     */
    private String title;
    /**
     * 样式集合，可以通过@ExcelCellStyle重复注解得到
     */
    private Map<String,CellStyle> cellStyleMap;
    /**
     * 是否启用样式的断言器
     */
    private Function<T,R> function;

}
