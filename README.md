## Excel表格转换工具包
### 用到的依赖：fastjson、poi-tl、lombok、spring-bean（只用到了字符串工具）
```xml
<!-- 这里面只用到了StringUtils做字符串判空其它没啥用可以自己修改源码 -->
<dependency>
    <groupId>org.springframework</groupId>
    <artifactId>spring-beans</artifactId>
    <version>5.3.7</version>
</dependency>
<!-- 操作excel的依赖工具包 -->
<dependency>
    <groupId>com.deepoove</groupId>
    <artifactId>poi-tl</artifactId>
    <version>1.9.1</version>
</dependency>
<!-- fastjson工具包 -->
<dependency>
    <groupId>com.alibaba</groupId>
    <artifactId>fastjson</artifactId>
    <version>1.2.76</version>
</dependency>
<!-- lombok 工具 -->
<dependency>
    <groupId>org.projectlombok</groupId>
    <artifactId>lombok</artifactId>
    <version>1.18.20</version>
    <scope>provided</scope>
</dependency>
```

### 注意
导入的excel表格的单元标题顺序可以变不影响最终结果，因为就是根据标题来确定位置的。只要这个单元格标题和对于的列是同一列即可


通过注解和工具类将，excel的数据并转换为List记录（后续数据库操作可以直接用MybatisPlus的批量插入即可）

导入示例在 top.yumbo.test.excel.importDemo.ImportExcelDemo 类的main方法
1.xlsx是测试用的表格
以下面的这张表为例：

<img src="https://img-blog.csdnimg.cn/20210523215535878.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>

## 通用的注解使用说明

### 第一步构建好实体在实体上加注解
#### 类注解，描述表头信息，表头高，表格的名称，excel模板的资源文件`@ExcelTableHeaderAnnotation(height = 4,tableName = "区域年度数据")`
height：以上图为例第5行才是我们需要导入的数据，表头也就是4行，这里的height就填4（excel是从1开始的，也就是标题是1、2、3、4这几行）
tableName：表示表格的名称如下图
<img src="https://img-blog.csdnimg.cn/20210523220150698.png"/>

#### 字段注解，与那个标题进行绑定，例如与标题为 "地区" 单元格绑定，表示这个字段的数据来自这个标题下 
`@ExcelCellBindAnnotation(title = "地区",width = 2,exception = "地区不存在")`
title：表是单元格属性列的标题
width：表示横向合并了多少个单元格
exception：自定义的异常消息
## 一、excel导入
#### 导入情景一、一个字段的数据由多个单元格合并而来
通过标题确定了这个字段和表格的下标index绑定，总共用width个单元格，作用就是将这几个单元格内容合并后赋值给该字段。
（因此这种情景建议字段类型尽量的设置为字符串）

#### 导入情景二、一个字段的数据来自单元格的部分内容
使用pattern实则正则表达式。
例如原文内容:"2020年第4季度"，其中2020要将它赋值给 字段year(Integer类型)，通过正则表达式"([0-9]{4})年"来得到2020这个字符串
正则会去最内部的那个"()"。

#### 注解示例：单元测试中也有注解好的实体类，没有加注解的字段就不做处理
```java
@Data
@ExcelTableHeaderAnnotation(height = 4,tableName = "区域年度数据")// 表头占4行
public class RegionYearETLSyncResponse {

    @ExcelCellBindAnnotation(title = "地区",width = 2,exception = "地区不存在")
    private String regionCode;

    /**
     * 年份
     */
    @ExcelCellBindAnnotation(title = "年份",exception = "年份格式不正确")
    private Integer year;
    /**
     * 对于不想返回的则不加注解即可,或者title为 ""
     */
    private BigDecimal calGeneralIncomeDivOutcome;

}
```
#### 使用方式

调用ExcelImportExportUtils.parseSheetToList(泛型,sheet表);即可返回List类型的数据

例如：
下载项目执行完整的测试代码在单元测试中的
[top.yumbo.test.excel.ExcelImportTest](https://github.com/1015770492/ExcelImportAndExport/blob/master/src/test/java/top/yumbo/test/excel/importDemo/ExcelImportDemo.java)
执行一下main即可完成转换
执行结果如下：
```bash
=====年度数据======
=======
RegionYearETLSyncResponse(regionCode=贵阳市,贵阳市, year=2020, regionGdp=200000000.0, regionGdpPerCapita=1000000.0, regionGdpRank=50.0, regionGdpPerCapitaRank=0.300, regionGdpGrowth=0.550, industryContributeGdp=0.300, generalUrbanizationRate=0.800, enableIncomePerCapita=1000000000.0, financeTotalIncome=1000000000.0, comprehensiveFinance=900000000.0, generalBudgetIncome=100000000.0, taxIncome=100000000.0, nonTaxIncome=100000000.0, governmentFundIncome=100000000.0, superiorSubsidyIncome=100000000.0, returnIncome=100000000.0, generalTransferIncome=100000000.0, specialTransferIncome=100000000.0, financeTotalOutcome=100000000.0, generalBudgetOutcome=100000000.0, generalBudgetIncomeRank=0.600, totalIncomeRank=0.100, calTaxDivGeneralIncome=null, calGeneralDivFinanceOutcome=null, calGeneralIncomeDivOutcome=null, superiorGovernmentGdp=100000000.0, superiorGovernmentTotalIncome=1000000000.0, calRegionDivSuperiorGdp=null, calFinanceIncomeRegionDivSuperior=null)
RegionYearETLSyncResponse(regionCode=贵阳市,南明区, year=2020, regionGdp=300000000.0, regionGdpPerCapita=1010000.0, regionGdpRank=51.0, regionGdpPerCapitaRank=0.330, regionGdpGrowth=0.560, industryContributeGdp=0.310, generalUrbanizationRate=0.810, enableIncomePerCapita=1100000000.0, financeTotalIncome=1100000000.0, comprehensiveFinance=1000000000.0, generalBudgetIncome=200000000.0, taxIncome=200000000.0, nonTaxIncome=200000000.0, governmentFundIncome=200000000.0, superiorSubsidyIncome=200000000.0, returnIncome=200000000.0, generalTransferIncome=200000000.0, specialTransferIncome=200000000.0, financeTotalOutcome=200000000.0, generalBudgetOutcome=200000000.0, generalBudgetIncomeRank=0.300, totalIncomeRank=0.110, calTaxDivGeneralIncome=null, calGeneralDivFinanceOutcome=null, calGeneralIncomeDivOutcome=null, superiorGovernmentGdp=200000000.0, superiorGovernmentTotalIncome=1100000000.0, calRegionDivSuperiorGdp=null, calFinanceIncomeRegionDivSuperior=null)
RegionYearETLSyncResponse(regionCode=贵阳市,云岩区, year=2020, regionGdp=400000000.0, regionGdpPerCapita=1020000.0, regionGdpRank=52.0, regionGdpPerCapitaRank=0.440, regionGdpGrowth=0.570, industryContributeGdp=0.320, generalUrbanizationRate=0.820, enableIncomePerCapita=1200000000.0, financeTotalIncome=1200000000.0, comprehensiveFinance=1100000000.0, generalBudgetIncome=300000000.0, taxIncome=300000000.0, nonTaxIncome=300000000.0, governmentFundIncome=300000000.0, superiorSubsidyIncome=300000000.0, returnIncome=300000000.0, generalTransferIncome=300000000.0, specialTransferIncome=300000000.0, financeTotalOutcome=300000000.0, generalBudgetOutcome=300000000.0, generalBudgetIncomeRank=0.550, totalIncomeRank=0.120, calTaxDivGeneralIncome=null, calGeneralDivFinanceOutcome=null, calGeneralIncomeDivOutcome=null, superiorGovernmentGdp=300000000.0, superiorGovernmentTotalIncome=1200000000.0, calRegionDivSuperiorGdp=null, calFinanceIncomeRegionDivSuperior=null)
RegionYearETLSyncResponse(regionCode=贵阳市,花溪区, year=2020, regionGdp=500000000.0, regionGdpPerCapita=1030000.0, regionGdpRank=53.0, regionGdpPerCapitaRank=0.200, regionGdpGrowth=0.580, industryContributeGdp=0.330, generalUrbanizationRate=0.830, enableIncomePerCapita=1300000000.0, financeTotalIncome=1300000000.0, comprehensiveFinance=1200000000.0, generalBudgetIncome=400000000.0, taxIncome=400000000.0, nonTaxIncome=400000000.0, governmentFundIncome=400000000.0, superiorSubsidyIncome=400000000.0, returnIncome=400000000.0, generalTransferIncome=400000000.0, specialTransferIncome=400000000.0, financeTotalOutcome=400000000.0, generalBudgetOutcome=400000000.0, generalBudgetIncomeRank=0.530, totalIncomeRank=0.130, calTaxDivGeneralIncome=null, calGeneralDivFinanceOutcome=null, calGeneralIncomeDivOutcome=null, superiorGovernmentGdp=400000000.0, superiorGovernmentTotalIncome=1300000000.0, calRegionDivSuperiorGdp=null, calFinanceIncomeRegionDivSuperior=null)
RegionYearETLSyncResponse(regionCode=贵阳市,白云区, year=2020, regionGdp=600000000.0, regionGdpPerCapita=1040000.0, regionGdpRank=54.0, regionGdpPerCapitaRank=0.400, regionGdpGrowth=0.590, industryContributeGdp=0.340, generalUrbanizationRate=0.840, enableIncomePerCapita=1400000000.0, financeTotalIncome=1400000000.0, comprehensiveFinance=1300000000.0, generalBudgetIncome=500000000.0, taxIncome=500000000.0, nonTaxIncome=500000000.0, governmentFundIncome=500000000.0, superiorSubsidyIncome=500000000.0, returnIncome=500000000.0, generalTransferIncome=500000000.0, specialTransferIncome=500000000.0, financeTotalOutcome=500000000.0, generalBudgetOutcome=500000000.0, generalBudgetIncomeRank=0.400, totalIncomeRank=0.140, calTaxDivGeneralIncome=null, calGeneralDivFinanceOutcome=null, calGeneralIncomeDivOutcome=null, superiorGovernmentGdp=500000000.0, superiorGovernmentTotalIncome=1400000000.0, calRegionDivSuperiorGdp=null, calFinanceIncomeRegionDivSuperior=null)
RegionYearETLSyncResponse(regionCode=贵阳市,观山湖区, year=2021, regionGdp=700000000.0, regionGdpPerCapita=1050000.0, regionGdpRank=55.0, regionGdpPerCapitaRank=0.410, regionGdpGrowth=0.600, industryContributeGdp=0.350, generalUrbanizationRate=0.850, enableIncomePerCapita=1500000000.0, financeTotalIncome=1500000000.0, comprehensiveFinance=1400000000.0, generalBudgetIncome=600000000.0, taxIncome=600000000.0, nonTaxIncome=600000000.0, governmentFundIncome=600000000.0, superiorSubsidyIncome=600000000.0, returnIncome=600000000.0, generalTransferIncome=600000000.0, specialTransferIncome=600000000.0, financeTotalOutcome=600000000.0, generalBudgetOutcome=600000000.0, generalBudgetIncomeRank=0.410, totalIncomeRank=0.150, calTaxDivGeneralIncome=null, calGeneralDivFinanceOutcome=null, calGeneralIncomeDivOutcome=null, superiorGovernmentGdp=600000000.0, superiorGovernmentTotalIncome=1500000000.0, calRegionDivSuperiorGdp=null, calFinanceIncomeRegionDivSuperior=null)
=====季度数据======
=======
RegionQuarterETLSyncResponse(year=2020, quarter=4, regionCode=贵阳市,贵阳市, breachNumber=2, breachTotalScale=100000000.0, riskNature=管理失误违约, riskVarieties=标准债券, regionDebtManage=强, calBondsHistoryCredit=是, repayCoordinated=强, cooperationCoordinated=强, sctDeployStatus=是)
RegionQuarterETLSyncResponse(year=2020, quarter=2, regionCode=贵阳市,南明区, breachNumber=3, breachTotalScale=200000000.0, riskNature=技术违约, riskVarieties=银行贷款, regionDebtManage=弱, calBondsHistoryCredit=否, repayCoordinated=弱, cooperationCoordinated=弱, sctDeployStatus=否)
RegionQuarterETLSyncResponse(year=2020, quarter=3, regionCode=贵阳市,云岩区, breachNumber=4, breachTotalScale=300000000.0, riskNature=实质违约, riskVarieties=非标集合产品, regionDebtManage=强, calBondsHistoryCredit=否, repayCoordinated=强, cooperationCoordinated=强, sctDeployStatus=是)
RegionQuarterETLSyncResponse(year=2020, quarter=2, regionCode=贵阳市,花溪区, breachNumber=5, breachTotalScale=400000000.0, riskNature=技术违约, riskVarieties=标准债券, regionDebtManage=弱, calBondsHistoryCredit=是, repayCoordinated=强, cooperationCoordinated=强, sctDeployStatus=否)
RegionQuarterETLSyncResponse(year=2020, quarter=1, regionCode=贵阳市,白云区, breachNumber=6, breachTotalScale=500000000.0, riskNature=管理失误违约, riskVarieties=单一产品, regionDebtManage=强, calBondsHistoryCredit=是, repayCoordinated=弱, cooperationCoordinated=弱, sctDeployStatus=是)
RegionQuarterETLSyncResponse(year=2020, quarter=1, regionCode=贵阳市,观山湖区, breachNumber=7, breachTotalScale=600000000.0, riskNature=实质违约, riskVarieties=单一产品, regionDebtManage=强, calBondsHistoryCredit=否, repayCoordinated=弱, cooperationCoordinated=弱, sctDeployStatus=是)
```

## 二、excel导出

### 导出情景一、一个字段拆分成多个单元格
例如：地区代码，通过数据库持久层框架，然后经过转换后。假设regionCode="贵阳市,南明区"这样的数据，
需要拆成两个单元格，分别是"州市"、"区县"
解决方式就是通过exportFormat="$0,$1"来将"贵阳市"、"南明区"拆出来，再根据index+width的方式填入对应单元格

$0表示被替换的第一个拆分出来的词例如"贵阳市",这里意味着你可以再添加内容，例如将exportFormat="贵州省$0,$1"
这样就变成了"贵州省贵阳市"、"南明区"，然后再填入单元格中

### 导出情景二、多个字段合并成一个单元格
例如：季度表的时间
只要给字段注入同一个标题title="时间"表示数据要填入这个标题下，填入的格式是exportFormat中定义的格式
例如: year=2020，quarter=4  需要合成 "2020年第4季度"
设置导出的格式exportFormat。例如 2020 要变成 "2020年" 那么就格式设置为 "$0年" 到时候这个"2020"就会替换这个"$0"
同理："4" 变成 "第4季度" exportFormat格式设置为"第$1季度"
"2020年" 和 "第4季度" 如果要拼成 "2020年第4季度" 那么还需要设置exportPriority， 默认值为0，注意它的值应当与你的$X
值上的数字相同，表示拼串的顺序，从0开始



```java
@Data
@ExcelTableHeaderAnnotation(height = 4, tableName = "区域季度数据")// 表头占4行
public class ExcelExportTemplateForQuarter {

    /**
     * 年份
     */
    @ExcelCellBindAnnotation(title = "时间", exportFormat = "$0年")
    private Integer year;

    /**
     * 季度，填写1到4的数字
     */
    @ExcelCellBindAnnotation(title = "时间", exportFormat = "第$1季", exportPriority = 1)
    private Integer quarter;

    /**
     * 地区代码，存储最末一级的地区代码就可以
     */
    @ExcelCellBindAnnotation(title = "地区", width = 2,exportSplit = ",", exportFormat = "$0,$1")
    private String regionCode;

    /**
     * 违约主体家数
     */
    @ExcelCellBindAnnotation(title = "违约主体家数", exception = "数值格式不正确")
    private Integer breachNumber;

    /**
     * 合计违约规模，单位亿
     */
    @ExcelCellBindAnnotation(title = "合计违约规模",size = BigDecimalUtils.ONE_HUNDRED_MILLION_STRING,exception = ExceptionMsg.INCORRECT_FORMAT_EXCEPTION)
    private BigDecimal breachTotalScale;

}
```

#### 导出成功的结果如下

<img src="https://img-blog.csdnimg.cn/20210530165048791.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70" />



#### 实现自定义样式高亮显示功能

##### 高亮行

设计原理结合java8的函数式接口，将其中样式的生效是否逻辑，通过断言型接口来判断，而断言的逻辑可以通过lambda表达式
实现逻辑的断言。

1、创建样式：
通过建造者模式设置单元格样式，代码示例如下：
```java
CellStyleEntity.builder().fontName("微软雅黑").fontSize(12).bgColor(9).build().getCellStyle(workbook)
```

2、创建断言器：
```java
final Function<ExcelExportTemplateForQuarter,Integer> functional = (one) -> {
    if (one.getRiskNature().equals("技术违约") && (one.getYear() % 2 == 0 || one.getQuarter() == 3)) {
        return 1;
    } else {
        return 0;
    }
};
```
##### 完整的代码示例：
```java
/**
 * 某些行高亮展示
 */
// 3种样式
final List<CellStyle> cellStyleList = Arrays.asList(
        CellStyleEntity.builder().fontName("微软雅黑").fontSize(12).bgColor(9).build().getCellStyle(workbook),
        CellStyleEntity.builder().fontSize(12).bgColor(9).foregroundColor(13).build().getCellStyle(workbook),
        CellStyleEntity.builder().fontName("微软雅黑").fontSize(12).bgColor(10).build().getCellStyle(workbook)
);
// 使用函数时接口返回样式的下标，然后就会将样式注入进去
ExcelImportExportUtils.filledListToSheetWithCellStyleByFunction(quarterList, cellStyleList, (one) -> {
    if (one.getRiskNature().equals("技术违约") && (one.getYear() % 2 == 0 || one.getQuarter() == 3)) {
        return 1;
    }else {
        return 0;
    }
}, workbook.getSheetAt(0));
```


导出结果示例：

<img src="https://img-blog.csdnimg.cn/20210604160604627.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70" />
使用案例：

[高亮行的示例代码](https://github.com/1015770492/ExcelImportAndExport/blob/master/src/test/java/top/yumbo/test/excel/exportDemo/ExcelExportDemo.java)

##### 高亮符合条件的单元格

高亮 市州 或 区县 内容为XX市的内容

结果示例：再上面的导出代码链接的 案例2中，取消注释即可

<img src="https://img-blog.csdnimg.cn/2021060416153964.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>


完整的代码
```java
final CellStyle cellStyle = CellStyleEntity.builder().fontName("微软雅黑").bold(true).fontSize(12).build().getCellStyle(workbook);
//
final CellStyle cellStyle3 = CellStyleEntity.builder().fontSize(12).fontColor(14).foregroundColor(13).build().getCellStyle(workbook);
final CellStyle cellStyle4 = CellStyleEntity.builder().fontSize(12).fontColor(10).bold(true).fontColor(14).foregroundColor(40).build().getCellStyle(workbook);
final TitlePredicateList<ExcelExportTemplateForQuarter> predicateList = new TitlePredicateList<>();
// 提供断言处理
Predicate<ExcelExportTemplateForQuarter> predicate = (e) -> {
    String regex = ".*市";// 高亮市
    final Pattern pattern = Pattern.compile(regex);
    final Matcher matcher = pattern.matcher(e.getRegionCode().split(",")[0]);
    if (matcher.matches()) {
        return true;
    }
    return false;
};
Predicate<ExcelExportTemplateForQuarter> predicate2 = (e) -> {
    String regex = ".*市";// 高亮市
    final Pattern pattern = Pattern.compile(regex);
    final Matcher matcher = pattern.matcher(e.getRegionCode().split(",")[1]);
    if (matcher.matches()) {
        return true;
    }
    return false;
};
Predicate<ExcelExportTemplateForQuarter> predicate3 = (e) -> {
    if (e.getRiskNature().equals("管理失误违约")) {
        return true;
    }
    return false;
};
// 高亮时间，第3季度的背景色设置为蓝色，字体红色加粗
Predicate<ExcelExportTemplateForQuarter> predicate4 = (e) -> {
    if (e.getQuarter() == 3) {
        return true;
    }
    return false;
};

final List<TitleCellStylePredicate<ExcelExportTemplateForQuarter>> titlePredicateList = predicateList
        .add("市州", cellStyle, predicate)
        .add("区县", cellStyle, predicate2)
        .add("风险性质", cellStyle3, predicate3)
        .add("时间", cellStyle4, predicate4)
        .getTitlePredicateList();
ExcelImportExportUtils.filledListToSheetWithCellStyleByBatchTitlePredicate(quarterList, titlePredicateList, sheet);
```

