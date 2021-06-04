package top.yumbo.excel.entity;


import org.apache.poi.ss.usermodel.CellStyle;

import java.util.ArrayList;
import java.util.List;
import java.util.function.Predicate;

/**
 * @author jinhua
 * @date 2021/6/4 11:07
 */
public class TitlePredicateList<T> {
    List<TitleCellStylePredicate<T>> list = new ArrayList<>();

    public TitlePredicateList<T> add(String title, CellStyle cellStyle, Predicate<T> predicate) {
        final TitleCellStylePredicate<T> onePredicate = new TitleCellStylePredicate<>(title, cellStyle, predicate);
        list.add(onePredicate);
        return this;
    }

    public List<TitleCellStylePredicate<T>> getTitlePredicateList() {
        return list;
    }
}
