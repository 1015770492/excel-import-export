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

### 1、导入excel数据并转换为List记录（后续数据库操作可以直接用MybatisPlus的批量插入即可）

导入示例在 top.yumbo.test.excel.ExcelImportDemo 类的main方法
1.xls是测试用的表格
以下面的这张表为例：

<img src="https://img-blog.csdnimg.cn/20210523215535878.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>

使用方式：
#### 第一步构建好实体在实体上加注解
##### 表头注解，注解在类名称上`@ExcelTableHeaderAnnotation(height = 4,tableName = "区域年度数据")`
height：以上图为例第5行才是我们需要导入的数据，表头也就是4行，这里的height就填4（excel是从1开始的，也就是标题是1、2、3、4这几行）
tableName：表示表格的名称如下图
<img src="https://img-blog.csdnimg.cn/20210523220150698.png"/>

##### 表的身体注解，注解在字段上`@ExcelCellBindAnnotation(title = "地区",width = 2,exception = "地区不存在")`
title：表是单元格属性列的标题
width：表示横向合并了多少个单元格
exception：自定义的异常消息

##### 注解示例：单元测试中也有注解好的实体类，没有加注解的字段就不做处理
```java
@Data
@ExcelTableHeaderAnnotation(height = 4,tableName = "区域年度数据")// 表头占4行
public class RegionYearETLSyncResponse {
    /**
     * 地区代码，存储最末一级的地区代码就可以
     */
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

```java
/**
 * 核心方法，传入泛型（带注解信息），sheet待解析的数据
 */
// 加了注解信息的实体类
List<RegionYearETLSyncResponse> regionYearETLSyncResponses = ExcelImportExportUtils.parseSheetToList(RegionYearETLSyncResponse.class, sheet);
```
下载项目执行完整的测试代码在单元测试中的top.yumbo.test.excel.ExcelImportTest的main方法，执行一下main即可看出效果
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