package top.yumbo.test.excel.importDemo;

import lombok.Data;
import top.yumbo.excel.annotation.ExcelTitleBind;
import top.yumbo.excel.annotation.ExcelTableHeader;

import java.io.Serializable;
import java.math.BigDecimal;
import java.time.LocalDate;

/**
 * @author Yujinhua
 * @Description
 * @date 2021/6/18 9:16
 */
@Data
@ExcelTableHeader( height = 2)
public class BIExcelResp implements Serializable {


    @ExcelTitleBind(title = "标准债券发行情况")
    private String w1;
    @ExcelTitleBind(title = "证券代码", nullable = true)
    private String w2;
    @ExcelTitleBind(title = "证券简称", nullable = true)
    private String w3;
    @ExcelTitleBind(title = "类别", nullable = true)
    private String w4;
    @ExcelTitleBind(title = "发行方式", nullable = true)
    private String w5;
    @ExcelTitleBind(title = "到期日期", nullable = true)
    private LocalDate w6;
    @ExcelTitleBind(title = "增信措施", nullable = true)
    private String w7;
    @ExcelTitleBind(title = "发行总额(万元)", size = "10000")
    private BigDecimal w8;
    @ExcelTitleBind(title = "可发行余额(万元)", size = "10000")
    private BigDecimal w9;
    @ExcelTitleBind(title = "主体评级")
    private String w10;
    @ExcelTitleBind(title = "债项评级")
    private String w11;
    @ExcelTitleBind(title = "评级日期")
    private LocalDate w12;
    @ExcelTitleBind(title = "评级机构名称")
    private String w13;
    @ExcelTitleBind(title = "评级展望")
    private String w14;

}
