package top.yumbo.excel.test.entity;

//import io.swagger.annotations.ApiModelProperty;

import lombok.Data;
import top.yumbo.excel.annotation.core.ExcelTableHeader;
import top.yumbo.excel.annotation.core.ExcelTitleBind;

/**
 * 投资项目导入类
 */
@Data
@ExcelTableHeader(height = 3, sheetName = "市属企业")
public class ImportRegularTitle {

    @ExcelTitleBind(title = "市属企业")
    private String w1;
    @ExcelTitleBind(title = ".*年计划")
    private String w2;
    @ExcelTitleBind(title = ".*月完成情况")
    private String w3;
    @ExcelTitleBind(title = ".*年计划")
    private String w4;
    @ExcelTitleBind(title = ".*月$")
    private String m1;
    @ExcelTitleBind(title = ".*月$")
    private String m2;
    @ExcelTitleBind(title = ".*月$")
    private String m3;
    @ExcelTitleBind(title = ".*月$")
    private String m4;
    @ExcelTitleBind(title = ".*月$")
    private String m5;
    @ExcelTitleBind(title = ".*月$")
    private String m6;
    @ExcelTitleBind(title = ".*月$")
    private String m7;
    @ExcelTitleBind(title = ".*月$")
    private String m8;
    @ExcelTitleBind(title = ".*月$")
    private String m9;
    @ExcelTitleBind(title = ".*月$")
    private String m10;
    @ExcelTitleBind(title = ".*月$")
    private String m11;

}
