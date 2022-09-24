package top.yumbo.excel.test.entity;

//import io.swagger.annotations.ApiModelProperty;

import lombok.Data;
import top.yumbo.excel.annotation.core.ExcelTableHeader;
import top.yumbo.excel.annotation.core.ExcelTitleBind;

@Data
@ExcelTableHeader(height = 4, sheetName = "县市区")
public class ImportRegularTitle_For_multiLevelTitle {

    @ExcelTitleBind(title = "县（市）区、开发（度假）区")
    private String A;
    // 2018
    @ExcelTitleBind(title = ".*年完成数$_绝对值")
    private String B;
    @ExcelTitleBind(title = ".*年完成数$_增速（%）")
    private String C;
    //2019
    @ExcelTitleBind(title = ".*年完成数$_绝对值")
    private String D;
    @ExcelTitleBind(title = ".*年完成数$_增速（%）")
    private String E;
    //2020
    @ExcelTitleBind(title = ".*年完成数$_绝对值")
    private String F;
    @ExcelTitleBind(title = ".*年完成数$_增速（%）")
    private String G;
    //2021
    @ExcelTitleBind(title = ".*年完成数$_绝对值")
    private String H;
    @ExcelTitleBind(title = ".*年完成数$_增速（%）")
    private String I;
    @ExcelTitleBind(title = "各县市区人代会目标_绝对值")
    private String J;
    @ExcelTitleBind(title = "各县市区人代会目标_增速（%）")
    private String K;

    @ExcelTitleBind(title = ".*年预期目标$_绝对值")
    private String L;
    @ExcelTitleBind(title = ".*年预期目标$_增速（%）")
    private String M;
    @ExcelTitleBind(title = ".*年1-2月完成数$_绝对值")
    private String N;
    @ExcelTitleBind(title = ".*年1-2月完成数$_增速（%）")
    private String O;

    @ExcelTitleBind(title = ".*年1-2月完成数$_绝对值")
    private String P;
    @ExcelTitleBind(title = ".*年1-2月完成数$_增速（%）")
    private String Q;

    /**
     * 2021年1-3月完成投资
     */
    @ExcelTitleBind(title = ".*完成投资$_绝对值")
    private String R;
    @ExcelTitleBind(title = ".*完成投资$_增速（%）")
    private String S;
    @ExcelTitleBind(title = ".*完成投资$_单月完成数")
    private String T;
    @ExcelTitleBind(title = ".*完成投资$_进度（%）")
    private String U;

    /**
     * 2021年1-4月完成投资
     */
    @ExcelTitleBind(title = ".*完成投资$_绝对值")
    private String V;
    @ExcelTitleBind(title = ".*完成投资$_增速（%）")
    private String W;
    @ExcelTitleBind(title = ".*完成投资$_单月完成数")
    private String X;
    @ExcelTitleBind(title = ".*完成投资$_进度（%）")
    private String Y;


}
