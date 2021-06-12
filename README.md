## excel-import-export能做的事情

1. 导入（复杂表单的导入，包括合并单元格的情况下也能进行导入，自带单位转换，通过设置size实现）
2. 导出（复杂表格的导出，包括一个字段多个单元格内容等情况的导出，并且实现高亮提示）

预计新增内容：

1. 字典样式会新增@MapEntry注解注入key，value字典数据（会在导入、导出自动转换字段和excel表格之间的数据）
2. （高亮行显示的前提下 或 默认样式下）部分单元格的字体，单元格样式的调整通过字段上注入@ExcelCellStyle，支持重复注解，后期使用java8的函数式接口返回这个字段具体使用那个样式。


底层的读取和写入都采用了forkjoin进行处理，性能上足够用，并且为了方便调整，额外提供了带threshold的方法
优点:
导入42M，79w条数据只需要71秒左右
导出文件可以实现高亮显示
数据量小的情况下可以随便使用，数据量比较大的情况下很可能会出现堆内存溢出。后续会堆内存进行优化。
如果内存足够可以适当的调整启动参数-Xmx8G，越大越好。

***

关于注解具体功能，查看内部注释即可。

<span style="color:red">
特别提醒：表头注解resource可以是http协议和https协议的excel模板文件，对于就版本的xls格式还需要注入type="xls"才可，否则因为兼容会报错。如果想要使用的是本地的文件作为模板也可以以path://开头，绝对路径则以 / 开头，相对路径直接文件夹开头。
</span>

导入功能不需要resource因为导入本身就会传一个excel，本身就是模板
导出功能：导出功能如果不通过注解方式提供resource，也可以用输入流的方式传入模板（需要传入是xls还是xlsx）

***



## Excel表格转换工具包

### 用到的依赖：fastjson、poi-tl、lombok、spring-bean（只用到了字符串工具）
## 引入依赖

[maven中央仓库地址（选择最新版本的文档会更新到最新版本）](https://mvnrepository.com/artifact/top.yumbo.excel/excel-import-export)

<span style="color:red">
特别提醒：由于国内很多使用的是aliyun的maven仓库依赖，阿里的仓库很可能同步的没有那么快，
           因此如果想要使用最新版本的，方式一、clone源码进行打包安装，然后引入坐标即可
           方式二、等阿里仓自动同步过去后引入坐标即可
</span>

以其中一个版本为例（选择最新版的）
将仓库地址改为下面的地址，以前的老版本仓库地址是：http://maven.aliyun.com/nexus/content/groups/public/
新仓库地址是：https://maven.aliyun.com/repository/central
经过测试老版本的仓库下载不了我同步在中央仓库的jar包，建议统一改成新版的ali镜仓地址
然后重启idea刷新一下maven依赖即可
全局修改maven仓库地址
```xml
<!--镜像仓库地址-->
<mirrors>
    <mirror>
      <id>alimaven</id>
      <name>aliyun maven</name>
      <url>https://maven.aliyun.com/repository/central</url>
      <mirrorOf>central</mirrorOf>
    </mirror>
</mirrors>
```

当前项目修改maven仓库地址

```xml
<repositories>
    <repository>
        <id>alimaven</id>
        <name>aliyun maven</name>
        <url>https://maven.aliyun.com/repository/central</url>
    </repository>
</repositories>
```

项目依赖地址：如果1.3.2不行就使用1.3.1 这两个版本内容一样。

```xml
<!-- https://mvnrepository.com/artifact/top.yumbo.excel/excel-import-export -->
<dependency>
    <groupId>top.yumbo.excel</groupId>
    <artifactId>excel-import-export</artifactId>
    <version>1.3.2</version>
</dependency>
```



### 注意
导入的excel表格的单元标题顺序可以变不影响最终结果，因为就是根据标题来确定位置的。只要这个单元格标题和对于的列是同一列即可


通过注解和工具类将，excel的数据并转换为List记录（后续数据库操作可以直接用MybatisPlus的批量插入即可）

导入示例在 [top.yumbo.test.excel.importDemo.ImportExcelDemo](https://github.com/1015770492/excel-import-export/blob/master/src/test/java/top/yumbo/test/excel/importDemo/ImportExcelDemo.java) 类的main方法
1.xlsx 是测试用的表格
以下面的这张表为例：

<img src="https://img-blog.csdnimg.cn/20210523215535878.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>

## 通用的注解使用说明

### 第一步构建好实体在实体上加注解
#### 类注解，描述表头信息，表头高，表格的名称，excel模板的资源文件`@ExcelTableHeader(height = 4,tableName = "区域年度数据")`
height：以上图为例第5行才是我们需要导入的数据，表头也就是4行，这里的height就填4（excel是从1开始的，也就是标题是1、2、3、4这几行）
tableName：表示表格的名称如下图
<img src="https://img-blog.csdnimg.cn/20210523220150698.png"/>

#### 字段注解，与那个标题进行绑定，例如与标题为 "地区" 单元格绑定，表示这个字段的数据来自这个标题下 
`@ExcelCellBind(title = "地区",width = 2,exception = "地区不存在")`
title：表是单元格属性列的标题
width：表示横向合并了多少个单元格
exception：自定义的异常消息

***


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
@ExcelTableHeader(height = 4,tableName = "区域年度数据")// 表头占4行
public class RegionYearETLSyncResponse {

    @ExcelCellBind(title = "地区",width = 2,exception = "地区不存在")
    private String regionCode;

    /**
     * 年份
     */
    @ExcelCellBind(title = "年份",exception = "年份格式不正确")
    private Integer year;
    /**
     * 对于不想返回的则不加注解即可,或者title为 ""
     */
    private BigDecimal calGeneralIncomeDivOutcome;

}
```

#### 使用方式

调用 即可返回List类型的数据
```java
ExcelImportExportUtils.importExcel(参数列表)
```
如下是导入导出的方法

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210613000905972.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)


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
/**
 * @author jinhua
 * @date 2021/5/28 14:28
 */
@Data
@ExcelTableHeader(height = 表高度int型, tableName = "表格名称", resource = "xls或者xlsx模板文件的类型",type="默认xlsx，对于xls文件填入xls（excel本身的版本兼容问题）")
// 表头占4行，使用了相对路径
public class ExportForQuarter {

    /**
     * 年份
     */
    @ExcelCellBind(title = "时间", importPattern = "([0-9]{4})年", exportFormat = "$0年")
    @ExcelCellStyle(id="1",表格的样式1,字体以及单元格样式的设置，具体看注解内部的功能)
    @ExcelCellStyle(id="2",表格的样式2，可以重复注解，为了实现部分样式的调整)
    private Integer year;

    /**
     * 季度，填写1到4的数字
     * 导入用到的则用importXXX命名，导出用exportXXX命名，其他则是通用的配置
     */
    @ExcelCellBind(title = "时间", importPattern = "([1-4]{1})季", exportFormat = "第$1季", exportPriority = 1)
    @ExcelCellStyle
    private Integer quarter;

    /**
     * 地区代码，存储最末一级的地区代码就可以
     */
    @ExcelCellBind(title = "地区", width = 2, exportSplit = ",", exportFormat = "$0,$1")
    @ExcelCellStyle(可以不加该注解，因为有默)
    private String regionCode;

    /**
     * 违约主体家数
     */
    @ExcelCellBind(title = "违约主体家数", exception = "默认异常：格式不正确，可以自定义异常提示")
    private Integer breachNumber;

    /**
     * 合计违约规模
     */
    @ExcelCellBind(title = "合计违约规模", size = "100000000")
    private BigDecimal breachTotalScale;

    /**
     * 风险性质 字典1260
     */
    @ExcelCellBind(title = "风险性质")
    private String riskNature;

    /**
     * 风险品种 字典1261
     */
    @ExcelCellBind(title = "风险品种")
    private String riskVarieties;

    /**
     * 区域偿债统筹管理能力 是否字典1022
     */
    @ExcelCellBind(title = "区域偿债统筹管理能力")
    private String regionDebtManage;

    /**
     * 区域内私募可转债历史信用记录 是否字典1022
     */
    @ExcelCellBind(title = "区域内私募可转债历史信用记录")
    private String calBondsHistoryCredit;

    /**
     * 还款可协调性 强弱字典1259
     */
    @ExcelCellBind(title = "还款可协调性")
    private String repayCoordinated;

    /**
     * 业务合作可协调性 强弱字典1259
     */
    @ExcelCellBind(title = "业务合作可协调性")
    private String cooperationCoordinated;

    /**
     * 数财通系统部署情况 是否字典1022
     */
    @ExcelCellBind(title = "数财通系统部署情况")
    private String sctDeployStatus;


}

```

#### 导出成功的结果如下

<img src="https://img-blog.csdnimg.cn/20210530165048791.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70" />



#### 实现自定义样式高亮显示功能

#### 一、导出的时候高亮行，根据每一行数据来自定义高亮行的样式
![在这里插入图片描述](https://img-blog.csdnimg.cn/20210607232159981.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)


例如按照第一季度的显示`黄色`、第二季度显示`玫瑰色`、第三季度显示`天蓝色`、第四季度显示`灰色`

示例代码
```java
/**
 * 得到List集合
 */
System.out.println("=====导入季度数据======");
String areaQuarter = "src/test/java/top/yumbo/test/excel/2.xlsx";
final List<ExportForQuarter> quarterList = ExcelImportExportUtils.importExcel(new FileInputStream(areaQuarter), ExportForQuarter.class, "xlsx");

/**
 * 将其导出
 */
if (quarterList != null) {
    quarterList.forEach(System.out::println);
    // 将数据导出到本地文件,如果要导出到web暴露出去只要传入输出流即可
    /**
     * 原样式导出
     */
    final Workbook workbook = ExcelImportExportUtils.exportExcel(quarterList, new FileOutputStream("D:/季度数据-原样式导出.xlsx"));
    /**
     * 高亮行方式导出
     */
    ExcelImportExportUtils.exportExcelRowHighLight(quarterList,
            new FileOutputStream("D:/季度数据-高亮行导出.xlsx"),
            (t) -> {
                if (t.getQuarter() == 1) {
                    return IndexedColors.YELLOW;
                } else if (t.getQuarter() == 2) {
                    return IndexedColors.ROSE;
                } else if (t.getQuarter() == 3) {
                    return IndexedColors.SKY_BLUE;
                } else if (t.getQuarter() == 4) {
                    return IndexedColors.GREY_25_PERCENT;
                }else {
                    return IndexedColors.WHITE;
                }
            });
}
```


使用案例：

[高亮行的示例代码](https://github.com/1015770492/ExcelImportAndExport/blob/master/src/test/java/top/yumbo/test/excel/exportDemo/ExcelExportDemo.java)

##### 高亮符合条件的单元格

