package top.yumbo.test.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import top.yumbo.excel.entity.CellStyleBuilder;

/**
 * @author jinhua
 * @date 2021/6/6 19:43
 */
public class CellStyleConstants {

    /**
     * 默认的样式
     */
    public static final CellStyle defaultStyle= CellStyleBuilder.builder()
            .fontName("微软雅黑")
            .foregroundColor(13)
            .build().getCellStyle();
    /**
     * 高亮
     */
    public static final CellStyle HIGH_LIGHT=CellStyleBuilder.builder()
            .fontName("微软雅黑")
            .foregroundColor(13)
            .build().getCellStyle();
}
