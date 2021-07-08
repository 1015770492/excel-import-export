## excel 导入的所有注解

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708160609772.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>

### 一、表头注解`@ExcelTableHeader`

用于记录数据行的起始位置，其中的height

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708165849251.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>

对应

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708165949864.png"/>

### 二、`@ExcelCellBind` 用于单元格和字段的绑定关系

#### 功能一、绑定单元格和字段，如果没有加的注解的字段不会进行处理

**title** 用于绑定单元格标题，根据标题进行绑定。

#### 功能二、支持绑定重复的单元格（后面有妙用！以字段为准）

例如：下面的多个字段有 **title** = `*房屋建筑面积单位`

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708164347136.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>

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

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708165312720.png" />

对应实体上的注解，分别表示截取时间单元格列下的部分内容：2021 和 4 。以此类推 2021 和 3 ...

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708165456628.png"/>

#### 功能六、读取多个相邻单元格，不相邻的单元格暂时没有做

**with** 用于获取多个单元格内容的合并内容

例如:

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708163930129.png"/>

可以用下面的内容来获取 贵阳市,南明市

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708163820479.png"/>

#### 功能七、进行空校验

**nullable** excel单元格是否可以为空，后面有一个更高级的空校验**`@CheckNullLogic`**逻辑空校验，意思是选择了某个值某些单元格必填，某些单元格必须置null

#### 功能八、字典替换规则

**replaceAll** 与**`@MapEntry`**结合使用，设置为**true**表示完全替换为value，为**false**表示将字段中的内容进行部分的替换

例如 "Abcabb" 如果`MapEntry(key="bb",value="cc")` 则会被替换为"Abcacc"。

如果是true 则需要进行强匹配 "Abcabb" 需要key="Abcabb" 才可以替换

#### 注解完整源码

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



### 二、`MapEntry` 用于字典转换，支持重复注解

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708162012616.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>

**注解案例：**

下面是一个单位的下拉框，有4个单位值

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708162332504.png"/>

注解上下面的信息后，其中的 *宗地面积单位 对应字段： `patriarchalAreaImUnit` ，值会被转换为 对应的value

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708162143450.png"/>



### 三、`AccountBigDecimalValue` 用于表格中的单位换算

单位的换算需要结合**@MapEntry**注解使用

可能你还需要额外的添加一个字段用于单位的映射，例如下面新增一个XXXSize的字段

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708172512527.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>



### 四、`CheckNullLogic` 用于逻辑校验空

注解

<img style="float:left" src="https://img-blog.csdnimg.cn/20210708172822387.png?x-oss-process=image/watermark,type_ZmFuZ3poZW5naGVpdGk,shadow_10,text_aHR0cHM6Ly9ibG9nLmNzZG4ubmV0L3FxXzQxODEzMjA4,size_16,color_FFFFFF,t_70"/>

**follow**表示需要对应的字段，**values**表示`follow`的值要是values中的一个，如果是其中一个，当前字段为null

会进行报错提示，提示的信息按照： followTitle值为XX 时，当前title的值不允许为null



