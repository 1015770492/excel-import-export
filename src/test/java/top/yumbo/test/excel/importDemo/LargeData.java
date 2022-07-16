package top.yumbo.test.excel.importDemo;

import lombok.Data;
import top.yumbo.excel.annotation.core.ExcelTableHeader;
import top.yumbo.excel.annotation.core.ExcelTitleBind;

@Data
@ExcelTableHeader(height = 2)
public class LargeData {

    @ExcelTitleBind(title = "编号")
    private String str1;

    @ExcelTitleBind(title = "名称")
    private String str2;

    @ExcelTitleBind(title = "城市")
    private String str3;

    @ExcelTitleBind(title = "状态")
    private String str4;

    @ExcelTitleBind(title = "备注")
    private String str5;

    @ExcelTitleBind(title = "备注一")
    private String str6;

    @ExcelTitleBind(title = "备注二")
    private String str7;

    @ExcelTitleBind(title = "备注三")
    private String str8;

    @ExcelTitleBind(title = "备注四")
    private String str9;

    @ExcelTitleBind(title = "备注五")
    private String str10;

    @ExcelTitleBind(title = "备注六")
    private String str11;

    @ExcelTitleBind(title = "备注七")
    private String str12;

    @ExcelTitleBind(title = "备注八")
    private String str13;

    @ExcelTitleBind(title = "备注一五")
    private String str14;

    @ExcelTitleBind(title = "备注二六")
    private String str15;

    @ExcelTitleBind(title = "备注三七")
    private String str16;

    @ExcelTitleBind(title = "备注四八")
    private String str17;

    @ExcelTitleBind(title = "备注五五")
    private String str18;

    @ExcelTitleBind(title = "备注六六")
    private String str19;

    @ExcelTitleBind(title = "备注七七")
    private String str20;

    @ExcelTitleBind(title = "备注八八")
    private String str21;

    @ExcelTitleBind(title = "备注九五")
    private String str22;

    @ExcelTitleBind(title = "备注六十")
    private String str23;

    @ExcelTitleBind(title = "备注七十")
    private String str24;

    @ExcelTitleBind(title = "备注八十")
    private String str25;
}