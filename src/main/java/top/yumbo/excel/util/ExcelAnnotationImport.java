package top.yumbo.excel.util;


import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.StringUtils;
import top.yumbo.excel.annotation.business.CheckNullLogic;
import top.yumbo.excel.annotation.business.ConvertBigDecimal;
import top.yumbo.excel.annotation.business.MapEntry;
import top.yumbo.excel.annotation.core.ExcelTableHeader;
import top.yumbo.excel.annotation.core.ExcelTitleBind;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author jinhua
 * @date 2021/5/21 21:51
 * <p>
 * 第二个版本的Excel导入工具，为了解决分支逻辑，
 * 为了解决选择某一个值后，某些字段不要填写问题
 * 整体的设计思路：
 * 1、导入：sheet->JsonObject->jsr303校验->收集起来并且添加进list
 * 转换中需要用到的信息通过注解实现，
 * 对于导入：通过注解-> 得到字段和表格关系
 */
public class ExcelAnnotationImport {

    /**
     * 类型map，如果后续还添加了其他类型则继续往下面添加
     */
    static HashMap<String, Class<?>> clazzMap = new HashMap<>();

    static {
        //如果新增了其他类型则继续put添加
        clazzMap.put(BigDecimal.class.getName(), BigDecimal.class);
        clazzMap.put(Date.class.getName(), Date.class);
        clazzMap.put(LocalDate.class.getName(), LocalDate.class);
        clazzMap.put(LocalDateTime.class.getName(), LocalDateTime.class);
        clazzMap.put(LocalTime.class.getName(), LocalTime.class);
        clazzMap.put(String.class.getName(), String.class);
        clazzMap.put(Byte.class.getName(), Byte.class);
        clazzMap.put(Short.class.getName(), Short.class);
        clazzMap.put(Character.class.getName(), Character.class);
        clazzMap.put(Integer.class.getName(), Integer.class);
        clazzMap.put(Float.class.getName(), Float.class);
        clazzMap.put(Double.class.getName(), Double.class);
        clazzMap.put(Long.class.getName(), Long.class);
    }

    //表头信息
    private enum TableEnum {
        TABLE_NAME, TABLE_HEADER, TABLE_HEADER_HEIGHT, TABLE_BODY, TITLE_MAP, REVERSE_MAP
    }

    // 单元格信息
    private enum CellEnum {
        TITLE_NAME, FIELD_NAME, FIELD_TYPE, SIZE, PATTERN, NULLABLE, WIDTH, EXCEPTION, COL, ROW, SPLIT, PRIORITY, FORMAT, MAP, REPLACE_ALL_OR_PART, SPLIT_REGEX
    }

    /**
     * 传入workbook不需要类型其他上面的
     *
     * @param sheet  要导入的sheet页
     * @param tClass ExcelResp类
     * @return ExcelResp类 List
     */
    public static <T> List<T> importExcel(Sheet sheet, Class<T> tClass) throws Exception {
        return sheetToList(sheet, tClass);
    }

    /**
     * 从输入流中获取woorkbook
     *
     * @param inputStream 输入流
     */
    public static Workbook getWorkBookByInputStream(InputStream inputStream) {
        // 1、如果输入流不为null，则从输入流中得到workbook
        if (inputStream != null) {
            try {
                return WorkbookFactory.create(inputStream);//可以读取xls格式或xlsx格式。
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    /**
     * 从输入流中获取一个sheet表格
     *
     * @param inputStream 输入流
     * @param sheetIdx    下标
     */
    public static Sheet getSheetByInputStream(InputStream inputStream, int sheetIdx) {
        Workbook workbook = getWorkBookByInputStream(inputStream);
        if (workbook != null) {
            return workbook.getSheetAt(sheetIdx);
        }
        return null;
    }

    /**
     * (导入)将sheet解析成List类型的数据
     * （注意这里只是将单元格内容转换为了实体，具体字段可能还不是正确的例如 区域码应该是是具体的编码而不是XX市XX区）
     *
     * @param tClass 传入的泛型
     * @param sheet  表单数据（带表头的）
     * @return 只是将单元格内容转化为List
     */
    private static <T> List<T> sheetToList(Sheet sheet, Class<T> tClass) throws Exception {
        if (sheet == null) {
            throw new NullPointerException("sheet不能为Null");
        }
        sheet.getWorkbook().setActiveSheet(0);
        final JSONObject importInfoByClazz = getImportInfoByClazz(sheet, tClass);
        final Integer tableHeight = getTableHeight(getTableHeaderDescInfo(importInfoByClazz));
        final JSONObject titleInfo = getExcelBodyDescInfo(importInfoByClazz);

        final int lastRowNum = sheet.getLastRowNum();

        return praseRowsToList(tableHeight, lastRowNum, sheet, titleInfo, importInfoByClazz.getJSONObject(TableEnum.TITLE_MAP.name()), tClass);
    }

    /**
     * 解析从start到end行的数据转换为List
     */
    private static <T> List<T> praseRowsToList(int start, int end, Sheet sheet, JSONObject fieldInfo, JSONObject titleMap, Class<T> tClass) throws Exception {
        // 从表头描述信息得到表头的高
        boolean flag = false;// 是否是异常行
        String message, lastExceptionMsg = "";// 异常消息
        List<T> list = new LinkedList<>();// 得到的所有数据结果
        ValidatorFactory vf = Validation.buildDefaultValidatorFactory();
        Validator validator = vf.getValidator();
        // 按行扫描excel表
        for (int i = start; i <= end; i++) {
            final Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            JSONObject oneRow = new JSONObject();// 一行数据
            oneRow.put(CellEnum.ROW.name(), i + 1);// 记录行号
            int length = fieldInfo.keySet().size();// 有多少个字段要进行处理
            int count = 0;// 记录异常空字段次数，如果与size相等说明是空行
            //将Row转换为JSONObject
            for (Object entry : fieldInfo.values()) {
                JSONObject fieldDesc = (JSONObject) entry;

                // 得到字段的索引位子
                Integer index = fieldDesc.getInteger(CellEnum.COL.name());
                if (index < 0) {
                    continue;
                }
                Integer width = fieldDesc.getInteger(CellEnum.WIDTH.name());// 得到宽度，如果宽度不为1则需要进行合并多个单元格的内容

                Boolean containsReplaceAll = fieldDesc.getBoolean(CellEnum.REPLACE_ALL_OR_PART.name());// 替换所有还是替换部分
                String splitRegex = fieldDesc.getString(CellEnum.SPLIT_REGEX.name());// 进行正则切割

                String fieldName = fieldDesc.getString(CellEnum.FIELD_NAME.name());// 字段名称
                String title = fieldDesc.getString(CellEnum.TITLE_NAME.name());// 标题名称
                String fieldType = fieldDesc.getString(CellEnum.FIELD_TYPE.name());// 字段类型
                String exception = fieldDesc.getString(CellEnum.EXCEPTION.name());// 转换异常返回的消息
                String size = fieldDesc.getString(CellEnum.SIZE.name());// 得到规模
                boolean nullable = fieldDesc.getBoolean(CellEnum.NULLABLE.name());
                final JSONObject map = fieldDesc.getJSONObject(CellEnum.MAP.name());// 字典


                String positionMessage = "第" + (i + 1) + "行,第" + (index + 1) + "列,标题：" + title;

                // 得到异常消息
                message = positionMessage + exception;

                // 获取合并的单元格值（合并后的结果，逗号分隔）
                String value = getMergeString(row, index, width, fieldType);
                if (map != null) {
                    if (StringUtils.hasText(splitRegex)) {
                        // 判断是否需要正则切割，有则将value进行切割处理
                        String[] split = value.split(splitRegex);
                        HashSet<String> strings = new HashSet<>();
                        StringBuilder stringBuilder = new StringBuilder();
                        for (int j = 0; j < split.length; j++) {
                            stringBuilder.append(replaceAllOrReplacePart(split[j], map, containsReplaceAll));
                            if (j + 1 < split.length) {
                                stringBuilder.append(splitRegex);
                            }
                        }
                        // 将处理完后的内容重新赋值给value
                        value = stringBuilder.toString();
                    } else {
                        // 没有内容就直接替换
                        value = replaceAllOrReplacePart(value, map, containsReplaceAll);
                    }
                }

                // 获取正则表达式，如果有正则，则进行正则截取value（相当于从单元格中取部分）
                String pattern = fieldDesc.getString(CellEnum.PATTERN.name());
                boolean hasText = StringUtils.hasText(value);
                Object castValue;
                // 默认字段不可以为空，如果注解过程设置为true则不抛异常
                if (!nullable) {
                    // 说明字段不可以为空
                    if (!hasText) {
                        // 字段不能为空结果为空，这个空字段异常计数+1。除非count==length，然后重新计数，否则就是一行异常数据
                        count++;// 不为空字段异常计数+1
                        lastExceptionMsg = message;
                        continue;
                    } else {
                        try {
                            // 单元格有内容,要么正常、要么异常直接抛不能返回null 直接中止
                            value = patternConvert(pattern, value);
                            castValue = cast(value, fieldType, size);
                        } catch (Exception e) {
                            throw new Exception(message + e.getMessage());
                        }
                    }

                } else {
                    // 字段可以为空 （要么正常 要么返回null不会抛异常）
                    length--;
                    try {
                        // 单元格内容无关紧要。要么正常转换，要么返回null
                        value = patternConvert(pattern, value);
                        castValue = cast(value, fieldType, size);
                    } catch (Exception e) {
                        throw new Exception(message + e.getMessage());
                    }
                }
                // 默认添加为null，只有正常才添加正常，否则中途抛异常直接中止

                oneRow.put(fieldName, castValue);// 添加数据
            }
            // 判断这行数据是否正常
            // 正常情况下count是等于length的，因为每个字段都需要处理
            if (count == 0) {
                checkLogic(oneRow, tClass, titleMap, i + 1);// 处理完后oneRow就是符合条件的数据
                updateBigDecimalValue(oneRow, tClass); // 更新单位
                T t = JSONObject.parseObject(oneRow.toJSONString(), tClass);
                // 进行jsr303校验
                Set<ConstraintViolation<T>> set = validator.validate(t);
                for (ConstraintViolation<T> constraintViolation : set) {
                    throw new Exception("第" + oneRow.getBigInteger(CellEnum.ROW.name()) + "行数据异常：" + constraintViolation.getMessage());
                }
                // 接着校验主从逻辑关系
                list.add(t);// 正常情况下添加一条数据
            } else if (count < length) {
                flag = true;// 需要抛异常，因为存在不合法数据
                break;// 非空行，并且遇到一行关键字段为null需要终止
            }
            // 空行继续扫描,或者正常
        }
        // 如果存在不合法数据抛异常
        if (flag) {
            throw new Exception(lastExceptionMsg);
        }
        return list;
    }

    /**
     * 根据单位进行更新值
     */
    private static <T> void updateBigDecimalValue(JSONObject oneRow, Class<T> tClass) {
        for (Field field : tClass.getDeclaredFields()) {
            ConvertBigDecimal accountBigDecimalValue = field.getDeclaredAnnotation(ConvertBigDecimal.class);
            if (accountBigDecimalValue != null) {
                String follow = accountBigDecimalValue.follow();

                String fieldName = field.getName();// 本身的字段名称
                BigDecimal bigDecimal = oneRow.getBigDecimal(follow);
                String size = oneRow.getString(fieldName);
                if (StringUtils.hasText(size)) {
                    String decimalFormat = accountBigDecimalValue.decimalFormat();
                    DecimalFormat df = new DecimalFormat();
                    df.applyPattern(decimalFormat);
                    String decimalValueStr = df.format(bigDecimal.multiply(new BigDecimal(size)).stripTrailingZeros());
                    BigDecimal newValue = new BigDecimal(decimalValueStr);
                    oneRow.put(follow, newValue);// 替换为新值
                }
            }
        }
    }

    /**
     * 得到映射结果
     * 大都情况下containsReplaceAll=true
     * 如果containsReplaceAll=false，需要注意替换部分，必须MapEntry中本身的key不能包含，否则就会替换错误的字典项
     */
    private static String replaceAllOrReplacePart(String value, JSONObject map, Boolean containsReplaceAll) {
        // 转换为字典项
        value = value.trim();// 去掉首尾多余空格等无实意符合
        if (containsReplaceAll) {
            // 是完全替换
            for (Map.Entry<String, Object> mapEntry : map.entrySet()) {
                if (value.equals(mapEntry.getKey())) {
                    value = mapEntry.getValue().toString();
                    break;
                }
            }

        } else {
            // 不是完全替换，只替换部分
            for (Map.Entry<String, Object> mapEntry : map.entrySet()) {
                if (value.contains(mapEntry.getKey())) {
                    value = value.replaceAll(mapEntry.getKey(), mapEntry.getValue().toString());
                    break;
                }
            }
        }
        return value;
    }

    /**
     * 校验非空逻辑
     */
    private static <T> void checkLogic(JSONObject data, Class<T> tClass, JSONObject titleMap, int row) {
        for (Field field : tClass.getDeclaredFields()) {
            CheckNullLogic annotation = field.getDeclaredAnnotation(CheckNullLogic.class);
            if (annotation != null) {
                // 校验follow的字段值是否符合values中的值
                String follow = annotation.follow();
                // 字典项的值
                String[] split = annotation.values();
                boolean needCheck = false;
                for (String value : split) {
                    if (StringUtils.hasText(follow)) {
                        Object obj = data.get(follow);
                        if (obj != null && value.equals(obj.toString())) {
                            // 需要进行校验，因为符合字典项
                            needCheck = true;
                            // 数据符合，接着校验当前field对应的值是否为null
                            Object fieldValue = data.get(field.getName());
                            if (fieldValue == null || !StringUtils.hasText(fieldValue.toString())) {
                                // 值为null或者""情况下，需要跑异常
                                // 得到follow字段上的标题，抛出提示信息。
                                ExcelTitleBind titleAnnotation = field.getDeclaredAnnotation(ExcelTitleBind.class);
                                titleMap.getJSONObject(follow).forEach((title, reverseMap) -> {
                                    if (reverseMap != null) {
                                        throw new RuntimeException("第" + row + "行，\"" + title + "\" 的值为:\"" + ((JSONObject) reverseMap).getString(value) + "\" 时，\"" + titleAnnotation.title() + "\" 值不能为空");
                                    }
                                });
                            } else {
                                // 不为null,说明校验通过了
                                break;
                            }
                        }
                    }

                }
                if (needCheck) {
                    // 符合字段项，并且字段不为null，通过处理下一个字段
                    continue;
                }
                // 不包含，本身这个数据不需要收集，故置null
                data.put(field.getName(), null);
            }
        }
    }

    /**
     * 获取表头的高度
     */
    private static Integer getTableHeight(JSONObject tableHeaderDesc) {
        return tableHeaderDesc.getInteger(TableEnum.TABLE_HEADER_HEIGHT.name());
    }

    /**
     * 从json中获取Excel身体部分数据
     */
    private static JSONObject getExcelBodyDescInfo(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(TableEnum.TABLE_BODY.name());
    }

    /**
     * 从json中获取Excel表头部分数据
     */
    private static JSONObject getTableHeaderDescInfo(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(TableEnum.TABLE_HEADER.name());
    }

    /**
     * 传入Sheet获取一个完整的表格描述信息，将INDEX更新
     *
     * @param excelDescData excel的描述信息
     * @param sheet         待解析的excel表格
     */
    private static JSONObject filledTitleIndexBySheet(JSONObject excelDescData, Sheet sheet) {
        JSONObject tableHeaderDesc = getTableHeaderDescInfo(excelDescData);
        JSONObject tableBodyDesc = getExcelBodyDescInfo(excelDescData);
        Integer height = getTableHeight(tableHeaderDesc);// 得到表头占据了那几行

        // 补充table每一项的index信息
        tableBodyDesc.forEach((fieldName, cellDesc) -> {
            // 扫描包含表头的那几行 记录需要记录的标题所在的索引列，填充INDEX
            for (int i = 0; i < height; i++) {
                Row row = sheet.getRow(i);// 得到第i行数据（在表头内）
                // 遍历这行所有单元格，然后得到表头进行比较找到标题和注解上的titleName相同的单元格
                row.forEach(cell -> {
                    // 得到单元格内容（统一为字符串类型）
                    String title = getStringCellValue(cell, String.class.getTypeName());

                    JSONObject cd = (JSONObject) cellDesc;
                    // 如果标题相同找到了这单元格，获取单元格下标存入
                    if (title.equals(cd.getString(CellEnum.TITLE_NAME.name()))) {
                        int columnIndex = cell.getColumnIndex();// 找到了则取出索引存入jsonObject
                        cd.put(CellEnum.COL.name(), columnIndex); // 补全描述信息
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
    private static <T> JSONObject getImportInfoByClazz(Sheet sheet, Class<T> tClass) {
        // 获取表格部分描述信息（根据泛型得到的）
        JSONObject partDescData = getImportDescriptionByClazz(tClass);
        System.out.println(partDescData);
        // 根据相同标题填充index
        return filledTitleIndexBySheet(partDescData, sheet);
    }

    /**
     * 根据注解类的注解信息获取excel表格的描述信息
     * 根据 @ExcelTableEnumHeaderAnnotation 得到表头占多少行，剩下的都是表的数据
     * 根据 @ExcelCellEnumBindAnnotation 得到字段和表格表头的映射关系以及宽度
     *
     * @param clazz 传入的泛型
     * @return 所有加了注解需要映射 标题和字段的Map集合
     */
    private static JSONObject getImportDescriptionByClazz(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();// 获取所有字段
        JSONObject excelDescData = new JSONObject();// excel的描述数据
        JSONObject tableBody = new JSONObject();// 表中主体数据信息
        JSONObject tableHeader = new JSONObject();// 表中主体数据信息

        // 收集字段-标题 的关系
        JSONObject titleMap = new JSONObject();
        // 1、先得到表头信息
        final ExcelTableHeader excelTableHeaderAnnotation = clazz.getAnnotation(ExcelTableHeader.class);
        if (excelTableHeaderAnnotation != null) {
            tableHeader.put(TableEnum.TABLE_NAME.name(), excelTableHeaderAnnotation.sheetName());// 表的名称
            tableHeader.put(TableEnum.TABLE_HEADER_HEIGHT.name(), excelTableHeaderAnnotation.height());// 表头的高度
            // 2、得到表的Body信息
            for (Field field : fields) {
                final ExcelTitleBind annotationTitle = field.getDeclaredAnnotation(ExcelTitleBind.class);

                if (annotationTitle != null) {// 找到自定义的注解
                    JSONObject cellDesc = new JSONObject();// 单元格描述信息
                    String title = annotationTitle.title();         // 获取标题，如果标题不存在则不进行处理
                    if (StringUtils.hasText(title)) {

                        // 获取字典映射
                        JSONObject mutiMap = getMapByMapEntries(field);
                        cellDesc.put(CellEnum.MAP.name(), mutiMap.get(CellEnum.MAP.name()));// 字典映射
                        JSONObject obj = new JSONObject();
                        obj.put(title, mutiMap.get(TableEnum.REVERSE_MAP.name()));// 字典反转
                        titleMap.put(field.getName(), obj);// 将反转Map存入titleMap中

                        cellDesc.put(CellEnum.SPLIT_REGEX.name(), annotationTitle.splitRegex()); // 正则切割符
                        cellDesc.put(CellEnum.REPLACE_ALL_OR_PART.name(), annotationTitle.replaceAll());// 是否包含替换所有,默认是替换所有

                        cellDesc.put(CellEnum.TITLE_NAME.name(), title);// 标题名称
                        cellDesc.put(CellEnum.FIELD_NAME.name(), field.getName());// 字段名称
                        cellDesc.put(CellEnum.FIELD_TYPE.name(), field.getType().getTypeName());// 字段的类型
                        cellDesc.put(CellEnum.WIDTH.name(), annotationTitle.width());// 单元格的宽度（宽度为2代表合并了2格单元格）
                        cellDesc.put(CellEnum.EXCEPTION.name(), annotationTitle.exception());// 校验如果失败返回的异常消息
                        cellDesc.put(CellEnum.COL.name(), annotationTitle.index());// 默认的索引位置
                        cellDesc.put(CellEnum.SIZE.name(), annotationTitle.size());// 规模,记录规模(亿元/万元)
                        cellDesc.put(CellEnum.PATTERN.name(), annotationTitle.importPattern());// 正则表达式
                        cellDesc.put(CellEnum.NULLABLE.name(), annotationTitle.nullable());// 是否可空
                        cellDesc.put(CellEnum.SPLIT.name(), annotationTitle.exportSplit());// 导出字段的拆分
                        cellDesc.put(CellEnum.FORMAT.name(), annotationTitle.exportFormat());// 导出的模板格式
                        cellDesc.put(CellEnum.PRIORITY.name(), annotationTitle.exportPriority());// 导出拼串的顺序

                        // 以字段名作为key
                        tableBody.put(field.getName(), cellDesc);// 存入这个标题名单元格的的描述信息，后面还需要补全INDEX
                    }
                }
            }
        }
        excelDescData.put(TableEnum.TABLE_HEADER.name(), tableHeader);// 将表头记录信息注入

        excelDescData.put(TableEnum.TITLE_MAP.name(), titleMap);// 将表头记录信息注入
        excelDescData.put(TableEnum.TABLE_BODY.name(), tableBody);// 将表的body记录信息注入
        return excelDescData;// 返回记录的所有信息
    }

    /**
     * 根据MapEntry注解得到字典映射，用于转换
     */
    private static JSONObject getMapByMapEntries(Field field) {
        final MapEntry[] mapEntries = field.getDeclaredAnnotationsByType(MapEntry.class);
        final JSONObject mutiMap = new JSONObject();
        final JSONObject map = new JSONObject();
        final JSONObject reverseMap = new JSONObject();

        for (MapEntry mapEntry : mapEntries) {
            if (mapEntry != null) {
                map.put(mapEntry.key(), mapEntry.value());
                reverseMap.put(mapEntry.value(), mapEntry.key());
            }
        }
        mutiMap.put(CellEnum.MAP.name(), map);
        mutiMap.put(TableEnum.REVERSE_MAP.name(), reverseMap);

        if (map.size() == 0 || reverseMap.size() == 0) {
            mutiMap.put(CellEnum.MAP.name(), null);
            mutiMap.put(TableEnum.REVERSE_MAP.name(), null);
        }
        return mutiMap;
    }

    /**
     * 单元格内容统一返回字符串类型的数据
     *
     * @param cell 单元格
     * @param type 字段类型
     */
    private static String getStringCellValue(Cell cell, String type) {
        String str = "";
        if (cell.getCellType() == CellType.STRING) {
            str = cell.getStringCellValue();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            if (clazzMap.get(type) == Date.class) {
                SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                final String format = formatter.format(cell.getDateCellValue());
                str += format;
            } else if (clazzMap.get(type) == LocalDateTime.class) {
                DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                String format = dtf.format(cell.getLocalDateTimeCellValue());
                str += format;
            } else if (clazzMap.get(type) == LocalDate.class) {
                DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd");
                String format = dtf.format(cell.getLocalDateTimeCellValue());
                str += format;
            } else if (clazzMap.get(type) == LocalTime.class) {
                DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm:ss");
                String format = dtf.format(cell.getLocalDateTimeCellValue());
                str += format;
            } else if (clazzMap.get(type) == String.class) {
                DataFormatter formatter = new DataFormatter();
                String formattedValue = formatter.formatCellValue(cell);
                str += formattedValue;
            } else {
                // 数值类型的
                str += cell.getNumericCellValue();
            }
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
    private static String getMergeString(Row row, Integer index, Integer width, String type) {
        // 合并单元格的处理方式 开始
        StringBuilder cellValue = new StringBuilder();
        for (int j = 0; j < width; j++) {
            String str = "";
            // 根据index得到单元格内容
            Cell cell = row.getCell(index + j);
            // 返回字符串类型的数据
            if (cell != null) {
                str = getStringCellValue(cell, type);
                if (str == null) {
                    str = "";
                }
            }

            if (j % 2 == 1) {
                if (StringUtils.hasText(cellValue.toString())) {
                    // 前面有文本才加逗号，前面没有文本就不加
                    cellValue.append(",");
                }
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
     * 根据类型的字符串得到返回类型
     */
    private static Object cast(String inputValue, String aClass, String size) throws Exception {
        return cast(inputValue, clazzMap.get(aClass), size);// 调用hashMap返回真正的类型
    }

    /**
     * 类型转换（返回转换的类型如果转换异常则抛出异常消息）
     *
     * @param inputValue 输入的待转换的字符串
     * @param aClass     转换成的类型
     * @throws Exception 转换异常抛出的异常消息
     */
    private static Object cast(String inputValue, Class<?> aClass, String size) throws
            Exception {
        Object obj = null;
        String value;
        if (StringUtils.hasText(inputValue)) {
            value = inputValue.trim();
        } else {
            return null;
        }
        try {
            if (aClass == BigDecimal.class) {
                obj = new BigDecimal(value).multiply(new BigDecimal(size)).stripTrailingZeros();
            } else if (aClass == String.class) {
                obj = value;
            } else if (aClass == Integer.class) {
                obj = new BigDecimal(value).intValue();
            } else if (aClass == Long.class) {
                obj = Long.parseLong(value.split("\\.")[0]);
            } else if (aClass == Double.class) {
                obj = Double.parseDouble(value);
            } else if (aClass == Short.class) {
                obj = new BigDecimal(value).shortValue();
            } else if (aClass == Character.class) {
                obj = value.charAt(0);
            } else if (aClass == Float.class) {
                obj = Float.parseFloat(value);
            } else if (aClass == Date.class) {
                obj = inputValue;
            } else if (aClass == LocalDateTime.class) {
                DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
                obj = LocalDateTime.parse(inputValue, dtf);
            } else if (aClass == LocalDate.class) {
                obj = LocalDate.parse(inputValue);
            } else if (aClass == LocalTime.class) {
                obj = LocalTime.parse(inputValue);
            }
        } catch (Exception e) {
            throw new Exception(" " + inputValue + " 格式不正确");
        }

        return obj;
    }

}
