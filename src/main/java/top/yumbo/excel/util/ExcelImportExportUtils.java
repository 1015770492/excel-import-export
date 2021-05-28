package top.yumbo.excel.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.StringUtils;
import top.yumbo.excel.annotation.ExcelCellBindAnnotation;
import top.yumbo.excel.annotation.ExcelTableHeaderAnnotation;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author jinhua
 * @date 2021/5/21 21:51
 */
public class ExcelImportExportUtils {
    public static final String INCORRECT_FORMAT_EXCEPTION = "格式不正确";
    public static final String CONVERT_EXCEPTION = "转换异常";
    public static final String NOT_BLANK_EXCEPTION = "不能为空";

    public enum ExcelTable {
        TABLE_NAME, TABLE_HEADER, TABLE_HEADER_HEIGHT, RESOURCE, TABLE_BODY;
    }

    public enum ExcelCell {
        TITLE_NAME, FIELD_NAME, FIELD_TYPE, SIZE, PATTERN, NULLABLE, WIDTH, EXCEPTION, INDEX, ROW;
    }


    /**
     * 批量保存（更新/插入）
     *
     * @param tList    传入的数据
     * @param iService 默认的批量保存
     * @param <T>      泛型
     */
//    public static <T> boolean defaultBatchSave(List<T> tList, IService<T> iService) {
//        return iService.saveBatch(tList);
//    }

    /**
     * 将sheet解析成List类型的数据（注意这里只是将单元格内容转换为了实体，具体字段可能还不是正确的例如 区域码应该是是具体的编码而不是XX市XX区）
     *
     * @param tClass 传入的泛型
     * @param sheet  表单数据（带表头的）
     * @return 只是将单元格内容转化为List
     */
    public static <T> List<T> parseSheetToList(Class<T> tClass, Sheet sheet) throws Exception {
        JSONArray jsonArray = parseSheetToJSONArray(tClass, sheet);
        return JSONArray.parseArray(jsonArray.toJSONString(), tClass);
    }

    /**
     * 将excel表转换为JSONArray
     *
     * @param tClass 注解模板类
     * @param sheet  传入的excel数据
     */
    public static <T> JSONArray parseSheetToJSONArray(Class<T> tClass, Sheet sheet) throws Exception {
        JSONObject fulledExcelDescData = getFulledExcelDescData(tClass, sheet);
        // 根据所有已知信息将excel转换为JsonArray数据
        return sheetToJSONArray(fulledExcelDescData, sheet);
    }

    /**
     * 存在模板的情况下
     * 将数据填充进入Excel表格
     */
    private static Sheet filledListToSheet(List<T> list, Sheet sheet) {
        // 得到完整的描述信息
        JSONObject fulledExcelDescData = getFulledExcelDescData(T.class, sheet);


        Workbook workbook = null;
        Sheet cloneSheet = workbook.cloneSheet(workbook.getSheetIndex(sheet));

        return null;
    }

    /**
     * 生成简单Excel表
     *
     * @param list  数据集
     * @param sheet excel表格
     */
    public static Workbook generateSimpleExcel(List<T> list, Sheet sheet) {

        // 根据反射获取表的所有信息
        JSONObject excelDescData = getExcelPartDescData(T.class);
        // 从所有数据中得到表头描述
        JSONObject excelHeaderDescData = getExcelHeaderDescData(excelDescData);
        // 从表头得到body数据，里面有title、index、field的绑定关系
        JSONObject excelBodyDescData = getExcelBodyDescData(excelDescData);

        Sheet filledSheet = filledListToSheet(list, sheet);
        return null;
    }

//    private static Sheet getResource(String resource) {
//        String proto = resource.split("://")[0];
//        if ("http".equals(proto) || "https".equals(proto)) {
//            // 网络下周好这个Excel模板文件
//        } else {
//            String path = getStringByPattern(resource, proto + "://(.*)");// 文件路径
//        }
//
//        return
//    }

    private static JSONArray sheetToJSONArray(JSONObject allExcelDescData, Sheet sheet) throws Exception {
        JSONArray result = new JSONArray();

        JSONObject tableHeader = allExcelDescData.getJSONObject(ExcelTable.TABLE_HEADER.name());// 表头描述信息
        JSONObject tableBody = allExcelDescData.getJSONObject(ExcelTable.TABLE_BODY.name());// 表的身体描述
        // 从表头描述信息得到表头的高
        Integer headerHeight = tableHeader.getInteger(ExcelTable.TABLE_HEADER_HEIGHT.name());
        final int lastRowNum = sheet.getLastRowNum();
        // 按行扫描excel表
        for (int i = headerHeight; i < lastRowNum; i++) {
            Row row = sheet.getRow(i);  // 得到第i行数据
            JSONObject oneRow = new JSONObject();// 一行数据
            int rowNum = i + 1;// 真正excel看到的行号
            oneRow.put(ExcelCell.ROW.name(), rowNum);// 记录行号
            int count = 0;// 记录空行次数
            //将Row转换为JSONObject
            for (Object entry : tableBody.values()) {
                JSONObject rowDesc = (JSONObject) entry;

                // 得到字段的索引位子
                Integer index = rowDesc.getInteger(ExcelCell.INDEX.name());
                if (index < 0) continue;
                Integer width = rowDesc.getInteger(ExcelCell.WIDTH.name());// 得到宽度，如果宽度不为1则需要进行合并多个单元格的内容

                String fieldName = rowDesc.getString(ExcelCell.FIELD_NAME.name());// 字段名称
                String title = rowDesc.getString(ExcelCell.TITLE_NAME.name());// 标题名称
                String fieldType = rowDesc.getString(ExcelCell.FIELD_TYPE.name());// 字段类型
                String exception = rowDesc.getString(ExcelCell.EXCEPTION.name());// 转换异常返回的消息
                String size = rowDesc.getString(ExcelCell.SIZE.name());// 得到规模
                boolean nullable = rowDesc.getBoolean(ExcelCell.NULLABLE.name());
                String positionMessage = "第" + rowNum + "行的,第" + (index + 1) + "列 ---标题：" + title + " -- ";
                //String nullExceptionMessage = positionMessage + NOT_BLANK_EXCEPTION;
                String castExceptionMessage = positionMessage + INCORRECT_FORMAT_EXCEPTION;
                // 返回的异常消息
                String message = positionMessage + exception;

                // 获取合并的单元格值（合并后的结果，逗号分隔）
                String value = getMergeString(row, index, width);

                // 获取正则表达式，如果有正则，则进行正则截取value（相当于从单元格中取部分）
                String pattern = rowDesc.getString(ExcelCell.PATTERN.name());
                boolean hasText = StringUtils.hasText(value);
                Object castValue = null;
                // 默认字段不可以为空，如果注解过程设置为true则不抛异常
                if (!nullable) {
                    // 说明字段不可以为空
                    if (!hasText) {
                        // 说明单元格内容为空需要抛异常
                        count++;
                        break;
                    } else {
                        try {
                            // 单元格有内容,要么正常、要么异常直接抛不能返回null 直接中止
                            value = patternConvert(pattern, value);
                            castValue = cast(value, fieldType, castExceptionMessage, size);
                        } catch (Exception e) {
                            throw new Exception(message);
                        }
                    }

                } else {
                    // 字段可以为空 （要么正常 要么返回null不会抛异常）
                    try {
                        // 单元格内容无关紧要。要么正常转换，要么返回null
                        value = patternConvert(pattern, value);
                        castValue = cast(value, fieldType, castExceptionMessage, size);
                    } catch (Exception e) {
                        //castValue=null;// 本来初始值就是null
                    }
                }
                // 默认添加为null，只有正常才添加正常，否则中途抛异常直接中止
                oneRow.put(fieldName, castValue);// 添加数据
            }

            if (count >= 1) {
                break;// 遇到一行关键字段为null需要终止
            }
            result.add(oneRow);// 添加一条数据
        }
        return result;
    }


    /**
     * 返回Excel主体数据Body的描述信息
     */
    public static JSONObject getExcelBodyDescData(Class<T> clazz) {
        JSONObject partDescData = getExcelPartDescData(clazz);
        return getExcelBodyDescData(partDescData);
    }

    /**
     * 从json中获取Excel身体部分数据
     */
    private static JSONObject getExcelBodyDescData(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(ExcelTable.TABLE_BODY.name());
    }

    /**
     * 返回Excel头部Header的描述信息
     */
    public static JSONObject getExcelHeaderDescData(Class<T> clazz) {
        JSONObject partDescData = getExcelPartDescData(clazz);
        return getExcelHeaderDescData(partDescData);
    }

    /**
     * 从json中获取Excel表头部分数据
     */
    private static JSONObject getExcelHeaderDescData(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(ExcelTable.TABLE_HEADER.name());
    }

    /**
     * 传入Sheet获取一个完整的表格描述信息，将INDEX更新
     *
     * @param excelDescData excel的描述信息
     * @param sheet         待解析的excel表格
     */
    private static JSONObject parseSheetAndFillTitleIndex(JSONObject excelDescData, Sheet sheet) {
        JSONObject tableHeaderDesc = excelDescData.getJSONObject(ExcelTable.TABLE_HEADER.name());
        JSONObject tableBodyDesc = excelDescData.getJSONObject(ExcelTable.TABLE_BODY.name());

        // 补充table每一项的index信息
        tableBodyDesc.forEach((fieldName, cellDesc) -> {
            Integer height = tableHeaderDesc.getInteger(ExcelTable.TABLE_HEADER_HEIGHT.name());// 得到表头占据了那几行
            // 扫描包含表头的那几行 记录需要记录的标题所在的索引列，填充INDEX
            for (int i = 0; i < height; i++) {
                Row row = sheet.getRow(i);// 得到第i行数据（在表头内）
                // 遍历这行所有单元格，然后得到表头进行比较找到标题和注解上的titleName相同的单元格
                row.forEach(cell -> {
                    // 得到单元格内容（统一为字符串类型）
                    String title = getStringCellValue(cell);

                    JSONObject cd = (JSONObject) cellDesc;
                    // 如果标题相同找到了这单元格，获取单元格下标存入
                    if (title.equals(cd.getString(ExcelCell.TITLE_NAME.name()))) {
                        int columnIndex = cell.getColumnIndex();// 找到了则取出索引存入jsonObject
                        cd.put(ExcelCell.INDEX.name(), columnIndex); // 补全描述信息
                    }
                });
            }

        });

        return excelDescData;
    }

    /**
     * 获取完整的Excel描述信息
     *
     * @param tClass 模板类
     * @param sheet  Excel
     * @param <T>    泛型
     */
    private static <T> JSONObject getFulledExcelDescData(Class<T> tClass, Sheet sheet) {
        // 获取表格部分描述信息（根据泛型得到的）
        JSONObject partDescData = getExcelPartDescData(tClass);
        // 根据相同标题填充index
        return parseSheetAndFillTitleIndex(partDescData, sheet);
    }

    /**
     * 获取excel表格的描述信息
     * 根据 @ExcelTableHeaderAnnotation 得到表头占多少行，剩下的都是表的数据
     * 根据 @ExcelCellBindAnnotation 得到字段和表格表头的映射关系以及宽度
     *
     * @param clazz 传入的泛型
     * @return 所有加了注解需要映射 标题和字段的Map集合
     */
    private static JSONObject getExcelPartDescData(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();// 获取所有字段
        JSONObject excelDescData = new JSONObject();// excel的描述数据
        JSONObject tableBody = new JSONObject();// 表中主体数据信息
        JSONObject tableHeader = new JSONObject();// 表中主体数据信息

        // 1、先得到表头信息
        Annotation clazzAnnotation = clazz.getAnnotation(ExcelTableHeaderAnnotation.class);
        if (clazzAnnotation != null) {
            ExcelTableHeaderAnnotation tableHeaderAnnotation = (ExcelTableHeaderAnnotation) clazzAnnotation;
            tableHeader.put(ExcelTable.TABLE_NAME.name(), tableHeaderAnnotation.tableName());// 表的名称
            tableHeader.put(ExcelTable.TABLE_HEADER_HEIGHT.name(), tableHeaderAnnotation.height());// 表头的高度
            tableHeader.put(ExcelTable.RESOURCE.name(), tableHeaderAnnotation.resource());// 模板excel的访问路径

            // 2、得到表的Body信息
            for (Field field : fields) {
                Annotation[] annotations = field.getDeclaredAnnotations();// 获取字段上所有注解
                for (Annotation annotation : annotations) {
                    if (annotation instanceof ExcelCellBindAnnotation) {// 找到自定义的注解
                        ExcelCellBindAnnotation annotationTitle = (ExcelCellBindAnnotation) annotation;// 进行强转
                        JSONObject cellDesc = new JSONObject();// 单元格描述信息
                        String title = annotationTitle.title();         // 获取标题，如果标题不存在则不进行处理
                        if (StringUtils.hasText(title)) {

                            int width = annotationTitle.width();// 获取注解的单元格宽度
                            String exception = annotationTitle.exception();// 获取异常信息
                            int index = annotationTitle.index();// 得到默认索引位置（默认-1，对于导出功能有奇效，导入可以不传）
                            int row = annotationTitle.row();// 标题所在的位置

                            cellDesc.put(ExcelCell.TITLE_NAME.name(), title);// 标题名称
                            cellDesc.put(ExcelCell.FIELD_NAME.name(), field.getName());// 字段名称
                            cellDesc.put(ExcelCell.FIELD_TYPE.name(), field.getType().getTypeName());// 字段的类型
                            cellDesc.put(ExcelCell.WIDTH.name(), width);// 单元格的宽度（宽度为2代表合并了2格单元格）
                            cellDesc.put(ExcelCell.EXCEPTION.name(), exception);// 校验如果失败返回的异常消息
                            cellDesc.put(ExcelCell.INDEX.name(), index);// 默认的索引位置
                            cellDesc.put(ExcelCell.ROW.name(), row);// 表示这行标题在那一行
                            cellDesc.put(ExcelCell.SIZE.name(), annotationTitle.size());// 规模,记录规模(亿元/万元)
                            cellDesc.put(ExcelCell.PATTERN.name(), annotationTitle.pattern());// 正则表达式
                            cellDesc.put(ExcelCell.NULLABLE.name(), annotationTitle.nullable());// 是否可空


                            // 以字段名作为key
                            tableBody.put(field.getName(), cellDesc);// 存入这个标题名单元格的的描述信息，后面还需要补全INDEX
                        }
                    }
                }
            }
        }
        excelDescData.put(ExcelTable.TABLE_HEADER.name(), tableHeader);// 将表头记录信息注入
        excelDescData.put(ExcelTable.TABLE_BODY.name(), tableBody);// 将表的body记录信息注入
        return excelDescData;// 返回记录的所有信息
    }


    /**
     * 单元格内容统一返回字符串类型的数据
     */
    private static String getStringCellValue(Cell cell) {
        String str = "";
        if (cell.getCellType() == CellType.STRING) {
            str = cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            str += cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.BOOLEAN) {
            str += cell.getBooleanCellValue();
        } else if (cell.getCellType() == CellType.BLANK) {
            str = "";
        }
        return str;
    }


    /**
     * 正则转换
     */
    private static String patternConvert(String pattern, String value) {
        if (StringUtils.hasText(pattern)) {
            // 如果存在正则，则单元格内容根据正则进行截取
            value = getStringByPattern(value, pattern);
        }
        return value;
    }

    /**
     * 获取单元格内容（逗号分隔）
     *
     * @param row   被取出的行
     * @param index 索引位置
     * @param width 索引位置+width确定取那几列
     * @return 返回合并单元格的内容（单个的则传width=1即可）
     */
    private static String getMergeString(Row row, Integer index, Integer width) {
        // 合并单元格的处理方式 开始
        StringBuilder cellValue = new StringBuilder();
        for (int j = 0; j < width; j++) {
            String str = "";
            // 根据index得到单元格内容
            Cell cell = row.getCell(index + j);
            // 返回字符串类型的数据
            if (cell != null) {
                str = getStringCellValue(cell);
                if (str == null) {
                    str = "";
                }
            }

            if (j % 2 == 1) {
                cellValue.append(",");
            }
            cellValue.append(str);
        }
        // 合并单元格处理结束
        return cellValue.toString();//得到单元格内容（合并后的）
    }

    /**
     * 根据正则获取内容（处理嵌套的正则并且得到最内部的字符串）
     *
     * @param inputString   输入的字符串
     * @param patternString 正则表达式字符串
     */
    private static String getStringByPattern(String inputString, String patternString) {
        String outputString = "";
        Pattern pattern = Pattern.compile(patternString);
        Matcher matcher = pattern.matcher(inputString);
        while (matcher.find()) {
            // 匹配最内部的那个正则匹配
            outputString = matcher.group(matcher.groupCount());
        }
        return outputString;
    }

    /**
     * 类型map，如果后续还添加了其他类型则继续往下面添加
     */
    static HashMap<String, Class<?>> clazzMap = new HashMap<>();

    static {
        //如果新增了其他类型则继续put添加
        clazzMap.put(BigDecimal.class.getName(), BigDecimal.class);
        clazzMap.put(String.class.getName(), String.class);
        clazzMap.put(Byte.class.getName(), Byte.class);
        clazzMap.put(Short.class.getName(), Short.class);
        clazzMap.put(Character.class.getName(), Character.class);
        clazzMap.put(Integer.class.getName(), Integer.class);
        clazzMap.put(Float.class.getName(), Float.class);
        clazzMap.put(Double.class.getName(), Double.class);
        clazzMap.put(Long.class.getName(), Long.class);
    }

    /**
     * 根据类型的字符串得到返回类型
     */
    private static Object cast(String inputValue, String aClass, String exception, String size) throws ClassCastException {
        return cast(inputValue, clazzMap.get(aClass), exception, size);// 调用hashMap返回真正的类型
    }

    /**
     * 类型转换（返回转换的类型如果转换异常则抛出异常消息）
     *
     * @param inputValue 输入的待转换的字符串
     * @param aClass     转换成的类型
     * @param exception  异常消息
     * @throws ClassCastException 转换异常抛出的异常消息
     */
    private static Object cast(String inputValue, Class<?> aClass, String exception, String size) throws ClassCastException {
        Object obj = null;
        String value;
        if (StringUtils.hasText(inputValue)) {
            value = inputValue.trim();
        } else {
            return null;
        }
        try {
            if (aClass == BigDecimal.class) {
                obj = new BigDecimal(value).multiply(new BigDecimal(size));// 乘以规模
            } else if (aClass == String.class) {
                obj = value;//直接返回字符串
            } else if (aClass == Integer.class) {
                obj = new Double(value).intValue();
            } else if (aClass == Long.class) {
                obj = Long.parseLong(value.split("\\.")[0]);// 小数点以后的不要
            } else if (aClass == Double.class) {
                obj = Double.parseDouble(value);
            } else if (aClass == Short.class) {
                obj = new Double(value).shortValue();
            } else if (aClass == Character.class) {
                obj = value.charAt(0);
            } else if (aClass == Float.class) {
                obj = Float.parseFloat(value);
            }
        } catch (Exception e) {
            System.out.println(inputValue);
            System.out.println(exception);
            throw new ClassCastException("类型转换异常，输入的文本内容：" + inputValue + "\n");
            //e.printStackTrace();
        }

        return obj;
    }

}
