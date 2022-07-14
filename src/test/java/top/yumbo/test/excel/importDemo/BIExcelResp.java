package top.yumbo.test.excel.importDemo;

import lombok.Data;
import top.yumbo.excel.annotation.core.TitleBind;
import top.yumbo.excel.annotation.core.TableHeader;

import java.io.Serializable;
import java.math.BigDecimal;
import java.time.LocalDate;

/**
 * @author Yujinhua
 * @Description
 * @date 2021/6/18 9:16
 */
@Data
@TableHeader(tableName = "标准债券发或拟发行情况（含评级信息）", height = 2)
public class BIExcelResp implements Serializable {


    @TitleBind(title = "标准债券发行情况")
    private String w1;
    @TitleBind(title = "证券代码", nullable = true)
    private String w2;
    @TitleBind(title = "证券简称", nullable = true)
    private String w3;
    @TitleBind(title = "类别", nullable = true)
    private String w4;
    @TitleBind(title = "发行方式", nullable = true)
    private String w5;
    @TitleBind(title = "到期日期", nullable = true)
    private LocalDate w6;
    @TitleBind(title = "增信措施", nullable = true)
    private String w7;
    @TitleBind(title = "发行总额(万元)", size = "10000")
    private BigDecimal w8;
    @TitleBind(title = "可发行余额(万元)", size = "10000")
    private BigDecimal w9;
    @TitleBind(title = "主体评级")
    private String w10;
    @TitleBind(title = "债项评级")
    private String w11;
    @TitleBind(title = "评级日期")
    private LocalDate w12;
    @TitleBind(title = "评级机构名称")
    private String w13;
    @TitleBind(title = "评级展望")
    private String w14;

}
