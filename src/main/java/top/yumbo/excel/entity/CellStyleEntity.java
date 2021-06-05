package top.yumbo.excel.entity;

import lombok.Builder;
import lombok.Data;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.StringUtils;

/**
 * @author jinhua
 * @date 2021/6/4 1:10
 */
@Builder
@Data
public class CellStyleEntity {
    private String fontName;
    private Integer fontSize;
    /**
     * 完整的颜色定义见：{@link org.apache.poi.ss.usermodel.IndexedColors}
     * 红10，白9，黑8，粉14，绿17，蓝12，
     * 灰22，金51，
     */
    private Integer fontColor;
    private Boolean bold;
    private Boolean locked;
    private Boolean hidden;
    private HorizontalAlignment textAlign;
    /**
     * 设置背景色
     * 完整的颜色定义见：{@link org.apache.poi.ss.usermodel.IndexedColors}
     * 默认白色，红10，白9，黑8，粉14，绿17，蓝12，
     * 灰22，金51，
     */
    private Integer bgColor;
    /**
     * 设置前景色（默认白色9）
     * 完整的颜色定义见：{@link org.apache.poi.ss.usermodel.IndexedColors}
     * 红10，白9，黑8，粉14，绿17，蓝12，
     * 灰22，金51，
     */
    private Integer foregroundColor;
    private Integer rotation;
    private VerticalAlignment verticalAlignment;
    private FillPatternType fillPatternType;
    private BorderStyle top;
    private BorderStyle bottom;
    private BorderStyle left;
    private BorderStyle right;
    private Boolean wrapText;
    private Boolean autoShrink;

    public CellStyle getCellStyle(Workbook wb){
        CellStyle cellStyle = wb.createCellStyle();
        Font font = wb.createFont();

        if (!StringUtils.hasText(fontName)) fontName = "微软雅黑";
        if (fontSize == null) fontSize = 11;
        if (fontColor == null) fontColor = 8;
        if (bgColor == null) bgColor = 9;
        if (bold == null) bold = false;
        if (rotation == null) rotation = 0;
        if (foregroundColor == null) foregroundColor = 9;
        if (locked == null) locked = false;
        if (hidden == null) hidden = false;
        if (wrapText == null) wrapText = false;
        if (autoShrink == null) autoShrink = false;
        if (textAlign == null) textAlign = HorizontalAlignment.CENTER;
        if (verticalAlignment == null) verticalAlignment = VerticalAlignment.CENTER;
        if (fillPatternType == null) fillPatternType = FillPatternType.SOLID_FOREGROUND;
        if (top == null) top = BorderStyle.THIN;
        if (bottom == null) bottom = BorderStyle.THIN;
        if (left == null) left = BorderStyle.THIN;
        if (right == null) right = BorderStyle.THIN;

        font.setFontName(fontName);// 字体
        font.setFontHeightInPoints(fontSize.shortValue());//设置字体大小
        font.setColor(fontColor.shortValue());
        font.setBold(bold);
        cellStyle.setFont(font);
        cellStyle.setLocked(locked);// 设置是否上锁，默认否
        cellStyle.setAlignment(textAlign);// 默认居中
        cellStyle.setRotation(rotation.shortValue());// 文字的旋转角度
        cellStyle.setVerticalAlignment(verticalAlignment);// 设置垂直方向的对齐
        cellStyle.setFillPattern(fillPatternType);// 设置填充前景色
        cellStyle.setBorderTop(top);// 设置边框类型，上
        cellStyle.setBorderBottom(bottom);// 下
        cellStyle.setBorderLeft(left);// 左
        cellStyle.setBorderRight(right);// 右
        cellStyle.setWrapText(wrapText);// 是否多行显示文本
        cellStyle.setShrinkToFit(autoShrink);// 如果文本太长，控制单元格是否应自动调整大小以缩小以适应
        cellStyle.setFillForegroundColor(foregroundColor.shortValue());// 设置前景色
        cellStyle.setFillBackgroundColor(bgColor.shortValue());// 设置背景色
        cellStyle.setHidden(hidden);//
        cellStyle.setFont(font);
        return cellStyle;
    }
}
