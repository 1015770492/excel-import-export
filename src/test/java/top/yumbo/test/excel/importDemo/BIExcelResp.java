package top.yumbo.test.excel.importDemo;

import com.alibaba.fastjson.annotation.JSONField;
import lombok.Data;
import top.yumbo.excel.annotation.ExcelCellBind;
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
@ExcelTableHeader(tableName = "标准债券发或拟发行情况（含评级信息）", height = 2)
public class BIExcelResp implements Serializable {


    @ExcelCellBind(title = "标准债券发行情况")
    private String w1;
    @ExcelCellBind(title = "证券代码", nullable = true)
    private String w2;
    @ExcelCellBind(title = "证券简称", nullable = true)
    private String w3;
    @ExcelCellBind(title = "类别", nullable = true)
    private String w4;
    @ExcelCellBind(title = "发行方式", nullable = true)
    private String w5;
    @ExcelCellBind(title = "到期日期", nullable = true)
    private LocalDate w6;
    @ExcelCellBind(title = "增信措施", nullable = true)
    private String w7;
    @ExcelCellBind(title = "发行总额(万元)", size = "10000")
    private BigDecimal w8;
    @ExcelCellBind(title = "可发行余额(万元)", size = "10000")
    private BigDecimal w9;
    @ExcelCellBind(title = "主体评级")
    private String w10;
    @ExcelCellBind(title = "债项评级")
    private String w11;
    @ExcelCellBind(title = "评级日期")
    private LocalDate w12;
    @ExcelCellBind(title = "评级机构名称")
    private String w13;
    @ExcelCellBind(title = "评级展望")
    private String w14;

}
