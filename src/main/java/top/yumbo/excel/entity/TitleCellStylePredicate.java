package top.yumbo.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.Predicate;

/**
 * @author jinhua
 * @date 2021/6/4 11:08
 */
@Data
@AllArgsConstructor
public class TitleCellStylePredicate<T> {

    private String title;
    private CellStyle cellStyle;
    private Predicate<T> predicate;


}
