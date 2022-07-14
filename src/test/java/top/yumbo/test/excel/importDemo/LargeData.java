package top.yumbo.test.excel.importDemo;

import lombok.Data;
import top.yumbo.excel.annotation.core.TableHeader;
import top.yumbo.excel.annotation.core.TitleBind;

@Data
@TableHeader(height = 2)
public class LargeData {

    @TitleBind(title = "编号")
    private String str1;

    @TitleBind(title = "名称")
    private String str2;

    @TitleBind(title = "城市")
    private String str3;

    @TitleBind(title = "状态")
    private String str4;

    @TitleBind(title = "备注")
    private String str5;

    @TitleBind(title = "备注一")
    private String str6;

    @TitleBind(title = "备注二")
    private String str7;

    @TitleBind(title = "备注三")
    private String str8;

    @TitleBind(title = "备注四")
    private String str9;

    @TitleBind(title = "备注五")
    private String str10;

    @TitleBind(title = "备注六")
    private String str11;

    @TitleBind(title = "备注七")
    private String str12;

    @TitleBind(title = "备注八")
    private String str13;

    @TitleBind(title = "备注一五")
    private String str14;

    @TitleBind(title = "备注二六")
    private String str15;

    @TitleBind(title = "备注三七")
    private String str16;

    @TitleBind(title = "备注四八")
    private String str17;

    @TitleBind(title = "备注五五")
    private String str18;

    @TitleBind(title = "备注六六")
    private String str19;

    @TitleBind(title = "备注七七")
    private String str20;

    @TitleBind(title = "备注八八")
    private String str21;

    @TitleBind(title = "备注九五")
    private String str22;

    @TitleBind(title = "备注六十")
    private String str23;

    @TitleBind(title = "备注七十")
    private String str24;

    @TitleBind(title = "备注八十")
    private String str25;
}