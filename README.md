# excel-import-export 
[在线文档](https://1015770492.github.io/excel-import-export/#/)

## 简单介绍

### 设计的起初原因

#### 导入的设计：

希望经过excel-import-export处理的数据是可以直接存入数据库，包括逻辑校验，字段空校验，jsr303校验。

自己曾测试过79w条数据，42M大小的excel导入，71秒的时间完成了所有数据从磁盘文件到java对象的转换

#### 导出的设计：

希望能快速的进行导出，并且可以带高亮的方式进行导出。

对于导出设计和导入相对应，对于合并的单元格处理，以及合并java字段内容，并且格式化导出也支持

举一个简单的例子：

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210709101819286.png)

对应字段

```java
@Data
// 表头占4行，height，同时使用resource设置模板文件位置
@ExcelTableHeader(height = 4, tableName = "区域季度数据", resource = "path://java/top/yumbo/test/excel/2_2.xlsx")
public class ExportForQuarter {

    // 根据正则截取单元格内容关于年份的值。其中exportFormat是导出excel填充到单元格的内容
    //	@Min(value = 2017,message = "最小年份是2017年") 支持jsr303校验注解
    //	@Max(value = 2021,message = "最大年份是2021年")
    @ExcelCellBind(title = "时间", exportFormat = "$0年")
    private Integer year;

    @ExcelCellBind(title = "时间", exportFormat = "第$1季度", exportPriority = 1)
    private Integer quarter;
}
```

2020会替换`exportFormat = "$0年"`中的`$0`，4会替换`exportFormat = "第$1季度"`的`$1`，

这两个字段会根据exportPriority进行拼串

就拼成了`2020年第4季度`填入时间这个单元格

## 快速开始

### 1、引入依赖

当前项目修改maven仓库地址

项目依赖地址：1.3.11

```xml
<repositories>
    <repository>
        <id>alimaven</id>
        <name>aliyun maven</name>
        <!-- 新版本的aliyun镜像仓库地址建议mirrors中也修改，
             如果已经改好了，则可以去掉这个repositories -->
        <url>https://maven.aliyun.com/repository/central</url>
    </repository>
</repositories>

<dependencies>

    <!-- https://mvnrepository.com/artifact/top.yumbo.excel/excel-import-export -->
    <dependency>
        <groupId>top.yumbo.excel</groupId>
        <artifactId>excel-import-export</artifactId>
        <version>1.3.11</version>
    </dependency>

</dependencies>
```



导入的excel表格的单元标题顺序可以变不影响最终结果，因为就是根据标题来确定位置的。只要这个单元格标题和对于的列是同一列即可


通过注解和工具类将，excel的数据并转换为List记录（后续数据库操作可以直接用MybatisPlus的批量插入即可）

导入示例在 [top.yumbo.test.excel.importDemo.ImportExcelDemo](https://github.com/1015770492/excel-import-export/blob/master/src/test/java/top/yumbo/test/excel/importDemo/ImportExcelDemo.java) 类的main方法
1.xlsx 是测试用的表格
以下面的这张表为例：

![](https://img-blog.csdnimg.cn/20210523215535878.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

### 2、表头注解`@ExcelTableHeader`
目的是得到表头占据了那几行，数据行应该从哪一行开始。

height：以上面年度数据为例，第5行是数据行，表头也就是4行，这里的height就填4
tableName：表的名称，相当于sheetName可以不填。



![](https://img-blog.csdnimg.cn/20210523220150698.png)

### 场景：

`@ExcelCellBind(title = "地区",width = 2,exception = "地区不存在")`
title：表是单元格属性列的标题
width：表示横向合并了多少个单元格
exception：自定义的异常消息

#### 一、excel导入

##### 导入情景一、合并多个单元格内容

合并：标题index+width个单元格`@ExcelCellBind(title = "地区",width = 2)`

单个单元格：直接`@ExcelCellBind(title = "年份")`

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210709091936476.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

对应下面的注解内容

```java
@Data
@ExcelTableHeader(height = 4)// 表头占4行
public class XXXForExcel {

    @ExcelCellBind(title = "地区",width = 2)
    private String regionCode;

    /**
     * 年份
     */
    @ExcelCellBind(title = "年份")
    private Integer year;
    /**
     * 对于不想返回的则不加注解即可,或者title为 ""
     */
    private BigDecimal calGeneralIncomeDivOutcome;

}
```



##### 导入情景二、截取单元格部分内容（正则截取）

例如单元格内容:"2020年第4季度"，其中2020要将它赋值给 字段year(Integer类型)，

可以通过正则表达式"([0-9]{4})年"来得到2020这个字符串

正则会取最内部的那个"()"。

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210709092207571.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

对应

```java
@Data
@ExcelTableHeader(height = 4, tableName = "区域季度数据")
public class ImportForQuarter {
    /**
     * 正则截取部分内容
     */
    @ExcelCellBind(title = "时间", importPattern = "([0-9]{4})年")
    private Integer year;
    @ExcelCellBind(title = "时间", importPattern = "([1-4]{1})季")
    private Integer quarter;
}
```

##### 导入情景三、自动单位换算

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210709092918254.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)



对应

```java
@Data
@ExcelTableHeader(height = 4, tableName = "区域季度数据")
public class ImportForQuarter {
    /**
     * 单位用size进行设置，例如表格上标注的单位是亿，这里的size就是下面的值。
     * 如果单位是%则填入字符串0.01即可以此类推
     */
    @ExcelCellBind(title = "合计违约规模",size = "100000000")
    private BigDecimal w5;
}
```

##### 导入情景四、不同单位的换算

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210709093339762.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

对应

```java
@Data
@ExcelTableHeader(height = 2)
public class PIMExcel {
    
    @ExcelCellBind(title = "*房屋建筑面积", nullable = true)
    private BigDecimal bAIM;

    @MapEntry(key = "公顷", value = "10000")
    @MapEntry(key = "平方公里", value = "1000000")
    @MapEntry(key = "平方米", value = "1")
    @MapEntry(key = "亩", value = "666.66667")
    @AccountBigDecimalValue(follow = "bAIM", decimalFormat = "#.##")
    @ExcelCellBind(title = "*房屋建筑面积单位")
    private String bAIMSize;
    
    // 这个用于存储，亩、公顷的信息。如果想要映射成字典可以加上@MapEntry
    @ExcelCellBind(title = "*房屋建筑面积单位")
    private String bAIMUnit;
}
```

@MapEntry 用于字典转换，这里换算成平方米

@AccountBigDecimalValue 用于单位的计算

例如：将12亩，换算成平方米

1. 用于存储换算后的值的字段类型用`BigDecimal`类型
2. 在单位字段上加`@AccountBigDecimalValue(follow = "bAIM", decimalFormat = "#.##")`其中decimalFormat用于格式化。#.## 表示最多保留2位小数，如果全是0则省略。#.00一定保留2位小数，不够就补0。以此类推

##### 导入情景五、逻辑空校验和置空

如下业务场景的逻辑校验

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210709095856212.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

对应着，需要注意一下，nullable需要职位true，表示可以为空，因为默认是不允许为空的。然后通过`@CheckNullLogic`进行逻辑校验

```java
@Data
@ExcelTableHeader(height = 2)
public class PIMExcel {
    
    @MapEntry(key = "土地使用权", value = "8001301")
    @MapEntry(key = "不动产", value = "8001302")
    @MapEntry(key = "在建工程", value = "8001303")
    @ExcelCellBind(title = "*抵押标的类型")
    private String mSTIm;
    
    @ExcelCellBind(title = "*房屋建筑面积", nullable = true)
    @CheckNullLogic(follow = "mSTIm", values = {"8001302", "8001303"})
    private BigDecimal bAIM;

	@CheckNullLogic(follow = "mSTIm", values = {"8001302", "8001303"})
    @ExcelCellBind(title = "*房屋建筑面积单位", nullable = true)
    private String bAIMUnit;
    
}
```

或者不用字典转换

```java
@Data
@ExcelTableHeader(height = 2)
public class PIMExcel {
    
    @ExcelCellBind(title = "*抵押标的类型")
    private String mSTIm;
    
    @ExcelCellBind(title = "*房屋建筑面积", nullable = true)
    @CheckNullLogic(follow = "mSTIm", values = {"不动产", "在建工程"})
    private BigDecimal bAIM;

	@CheckNullLogic(follow = "mSTIm", values = {"不动产", "在建工程"})
    @ExcelCellBind(title = "*房屋建筑面积单位", nullable = true)
    private String bAIMUnit;
    
}
```

用好字典转换和逻辑较空，处理完成后的数据是可以直接存数据库的。





#### 二、excel导出

##### 导出情景一、一个字段拆分成多个单元格

例如：地区代码，通过数据库持久层框架，然后经过转换后。假设regionCode="贵阳市,南明区"这样的数据，
需要拆成两个单元格，分别是"州市"、"区县"
解决方式就是通过exportFormat="$0,$1"来将"贵阳市"、"南明区"拆出来，再根据index+width的方式填入对应单元格

$0表示被替换的第一个拆分出来的词例如"贵阳市",这里意味着你可以再添加内容，例如将exportFormat="贵州省$0,$1"
这样就变成了"贵州省贵阳市"、"南明区"，然后再填入单元格中

##### 导出情景二、多个字段合并成一个单元格

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

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210530165048791.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)



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

高亮符合条件的单元格

## 全套注解

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708160609772.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

### 1、`@ExcelTableHeader`

用于记录数据行的起始位置，其中的height

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708165849251.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

对应

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708165949864.png)

### 2、`@ExcelCellBind` 

用于单元格和字段的绑定关系

#### 功能一、绑定单元格和字段，如果没有加的注解的字段不会进行处理

**title** 用于绑定单元格标题，根据标题进行绑定。

#### 功能二、支持绑定重复的单元格（后面有妙用！以字段为准）

例如：下面的多个字段有 **title** = `*房屋建筑面积单位`

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708164347136.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

#### 功能三、自动类型转换，以字段类型为准

类型转换问题

1. 日期类型：使用 LocalDate类型的

2. 数值类型：建议用BigDecimal类型的。当然也支持（Integer、Long、Short、BigDecimal、Float、Double）
3. 字符串类型：原封不动

#### 功能四、进行简单的单位换算

**size** 字段是数值类型的，并且size设置了单位值，会对字段值进行单位的换算。

例如：

万 对应 size="10000"，

% 对应 size="0.01"

并且字段是数值类型的即可进行单位的转换（Integer、Long、Short、BigDecimal、Float、Double）

#### 功能五、正则截取单元格部分内容

例如：时间的截取

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708165312720.png)

对应实体上的注解，分别表示截取时间单元格列下的部分内容：2021 和 4 。以此类推 2021 和 3 ...

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708165456628.png)

#### 功能六、读取多个相邻单元格，不相邻的单元格暂时没有做

**with** 用于获取多个单元格内容的合并内容

例如:

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708163930129.png)

可以用下面的内容来获取 贵阳市,南明市

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708163820479.png)

#### 功能七、空校验

**nullable** excel单元格是否可以为空，后面有一个更高级的空校验**`@CheckNullLogic`**逻辑空校验，意思是选择了某个值某些单元格必填，某些单元格必须置null

#### 功能八、字典替换规则

**replaceAll** 与**`@MapEntry`**结合使用，设置为**true**表示完全替换为value，为**false**表示将字段中的内容进行部分的替换

例如 "Abcabb" 如果`MapEntry(key="bb",value="cc")` 则会被替换为"Abcacc"。

如果是true 则需要进行强匹配 "Abcabb" 需要key="Abcabb" 才可以替换

#### 完整源码和注释

部分内容用于导出，和其它功能

```java
public @interface ExcelCellBind {
    /**
     * 绑定的标题名称，
     * 通过扫描单元格表头可以确定表头所在的索引列，然后在根据width就能确定单元格
     */
    String title() default "";
    /**
     * 单元格宽度，对于合并单元格的处理
     * 确定表格的位置采用： 下标（解析过程会得到下标） + 单元格的宽度
     * 这样就可以确定单元格的位子和占据的宽度
     */
    int width() default 1;
    /**
     * 注入的异常消息，为了校验单元格内容
     * 校验失败应该返回的消息提升
     */
    String exception() default "格式不正确";
    /**
     * 规模，对于BigDecimal类型的需要进行转换
     */
    String size() default "1";
    /**
     * 正则截取单元格部分内容，只需要部分其它内容丢掉
     * 一个单元格中的部分内容，例如 2020年2季度，只想单独取出年、季度这两个数字
     */
    String importPattern() default "";
    /**
     * 正则截取单元格内容，保留单元格内容，后面进行替换字典
     * 服务于replaceAllOrPart，如果使用了splitRegex，则会将内容切割进行replaceAllOrPart
     * 然后将将处理后的结果返回，然后再进行importPattern
     */
    String splitRegex() default "";
    /**
     * 包含字典key就完全替换为value
     * 例如：key=江西上饶, value=jx
     * replaceAll=true，那么就会被替换为jx。
     * 如果设置为false，只会替换字典部分内容，也就是变成：jx上饶
     */
    boolean replaceAll() default true;
    /**
     * 导出的字符串格式化填入，利用StringFormat.format进行字符串占位和替换
     */
    String exportFormat() default "";
    /**
     * 导出功能，该字段可能是多个单元格的内容（连续单元格），按照split拆分和填充。默认逗号
     */
    String exportSplit() default "";

    /**
     * 合并多个字段的顺序，多个字段构成一个标题，例如时间 年+季度
     */
    int exportPriority() default 0;
    /**
     * 默认不可以为空
     */
    boolean nullable() default false;
    /**
     * 单元格索引位置
     */
    int index() default -1;
}
```

### 3、`@MapEntry` 

用于字典转换，支持重复注解

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708162012616.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

**注解案例：**

下面是一个单位的下拉框，有4个单位值

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708162332504.png)

注解上下面的信息后，其中的 *宗地面积单位 对应字段： `patriarchalAreaImUnit` ，值会被转换为 对应的value

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708162143450.png)

### 4、`@AccountBigDecimalValue` 用于表格中的单位换算

单位的换算需要结合**@MapEntry**注解使用

可能你还需要额外的添加一个字段用于单位的映射，例如下面新增一个XXXSize的字段

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708172512527.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

### 5、`@CheckNullLogic` 

用于逻辑校验空

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210708172822387.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

1. **follow**表示需要对应的字段，

2. **values**表示`follow`的值要是values中的一个，如果是其中一个，当前字段为null

会进行报错提示，提示的信息按照： followTitle值为XX 时，当前title的值不允许为null

### 6、`@CheckValues`

字段值强校验

**values**字段值的数组，加上之后字段值必须为values数组中的值

**message**字段不符合强校验的情况下的异常消息提醒

### 7、`@ExcelCellStyle`

用于样式的设置，后续会加入，暂时没有做完整的设计

## 内置Utils

### CheckLogicUtils

用于逻辑空校验，支持jsr303校验。

实验方式：实体层加上jsr303校验，以及逻辑空校验注解`@CheckNullLogic`与Excel导入导出的逻辑校验相同，

只是符合其他场景下的任意实体的逻辑校验。

返回的对象是经过校验后的对象，注意返回的是一个新的对象！！

```java
request = CheckLogicUtils.checkNullLogicWithJSR303(request);
```

以前的写法

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210710071536984.png)

新的写法

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210710071649392.png)

### ExcelImportUtils2

单线程方式，专门为导入设计的工具。

结合配套注解使用，调用`ExcelImportUtils2.importExcel(Sheet sheet, Class<T> tClass)`方法即可返回List的实体数据

返回的数据与T类型相同（即返回的是加了注解信息的excel模板类）然后可以通过

`JSONObject.parseArray(JSON.toJSONString(返回的list),想要返回的实体.class)`或者

遍历list，然后加入到

`BeanUtils.copyProperties(excel模板实体类,想要返回的实体.class);`

### ExcelImportExportUtils

支持并发和单线程的导入和导出（可以带高亮）

具体的测试代码，看单元测试中案例

![在这里插入图片描述](https://img-blog.csdnimg.cn/20210710072333293.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70)

