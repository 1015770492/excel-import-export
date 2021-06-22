package top.yumbo.excel.entity;

import lombok.Data;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * @author jinhua
 * @date 2021/6/22 21:05
 * 标题建造器
 */
@Data
public class TitleBuilders {
    // 每一行的标题
    private List<List<TitleBuilder>> titleList= new ArrayList<>();;

    // 建造者模式得到一个自己
    public static TitleBuilders builder() {
        return new TitleBuilders();
    }
    /**
     * 添加一行标题
     */
    public TitleBuilders addOneRow(TitleBuilder... titleBuilders) {
        final List<TitleBuilder> list = Arrays.asList(titleBuilders);
        titleList.add(list);
        return this;
    }

    /**
     * 返回构建好的标题信息
     */
    public TitleBuilders build() {
        return this;
    }
}
