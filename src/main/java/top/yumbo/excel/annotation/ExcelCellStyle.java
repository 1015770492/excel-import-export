package top.yumbo.excel.annotation;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.*;

/**
 * @author jinhua
 * @date 2021/6/2 10:20
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Repeatable(ExcelCellStyles.class)
public @interface ExcelCellStyle {

    /**
     * 样式的id
     */
    String id() default "1";

    /**
     * 是否使隐藏样式（是否启用样式）
     */
    boolean hidden() default false;

    /**
     * 是否锁定单元格（可编辑/不可编辑）,默认不上锁（可编辑）
     */
    boolean locked() default false;

    /**
     * 默认微软雅黑
     */
    String fontName() default "微软雅黑";

    /**
     * 默认字体11号
     */
    int fontSize() default 11;

    /**
     * 字体加粗
     */
    boolean bold() default false;

    /**
     * 默认居中，单元格水平对齐方式的枚举值
     */
    HorizontalAlignment textAlign() default HorizontalAlignment.CENTER;

    /**
     * 垂直方向的对齐方式，默认居中
     */
    VerticalAlignment verticalAlign() default VerticalAlignment.CENTER;

    /**
     * 背景颜色   详情见：{@link org.apache.poi.ss.usermodel.IndexedColors}
     */
    int backgroundColor() default 9;

    /**
     * 文字旋转角度
     */
    int rotation() default 0;

    /**
     * 默认白色是9  {@link org.apache.poi.ss.usermodel.IndexedColors}
     */
    int foregroundColor() default 9;

    /**
     * 填充图案，钻石、细点等，默认不填充
     */
    FillPatternType fillPatternType() default FillPatternType.SOLID_FOREGROUND;

    /**
     * 控制单元格是否应自动调整大小以缩小以适应
     */
    boolean autoShrink() default false;

    /**
     * 多行显示文本内容
     */
    boolean wrapText() default false;

    /**
     * 上、下、左、右 边框样式
     */
    BorderStyle top() default BorderStyle.THIN;

    BorderStyle bottom() default BorderStyle.THIN;

    BorderStyle left() default BorderStyle.THIN;

    BorderStyle right() default BorderStyle.THIN;

}
