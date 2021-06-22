package top.yumbo.excel.entity;

import lombok.Data;

/**
 * @author jinhua
 * @date 2021/6/22 21:01
 */

@Data
public class TitleBuilder {
    // 下标
    private int index = 0;
    // 标题
    private String title;
    // 宽度
    private int width = 1;
    // 高度
    private int height = 1;

    public TitleBuilder index(int index) {
        this.index = index;
        return this;
    }

    public TitleBuilder title(String title) {
        this.title = title;
        return this;
    }

    public TitleBuilder width(int width) {
        this.width = width;
        return this;
    }

    public TitleBuilder height(int height) {
        this.height = height;
        return this;
    }

    public static TitleBuilder builder() {
        return new TitleBuilder();
    }

    public TitleBuilder build() {
        return this;
    }
}
