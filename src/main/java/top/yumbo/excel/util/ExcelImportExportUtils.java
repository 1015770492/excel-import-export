package top.yumbo.excel.util;


import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import lombok.Builder;
import lombok.Data;
import org.apache.commons.math3.exception.OutOfRangeException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;
import org.springframework.web.client.RestTemplate;
import top.yumbo.excel.annotation.business.ExcelCellStyle;
import top.yumbo.excel.annotation.core.ExcelTableHeader;
import top.yumbo.excel.annotation.core.ExcelTitleBind;
import top.yumbo.excel.consts.ExcelConstants;
import top.yumbo.excel.entity.CellStyleBuilder;
import top.yumbo.excel.entity.TitleBuilder;
import top.yumbo.excel.entity.TitleBuilders;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.ForkJoinPool;
import java.util.concurrent.RecursiveTask;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import static top.yumbo.excel.consts.ExcelConstants.intMap;

/**
 * @author jinhua
 * @date 2021/5/21 21:51
 * <p>
 * 整体的设计思路：
 * 1、导入：sheet->JsonObject->jsr303校验->收集起来并且添加进list
 * 2、导出：list->jsr303校验->填入数据（可以高亮行的填入）->sheet
 * 设计思路
 * 对于导入：通过注解-> 得到字段和表格关系
 * 对于导出：通过注解-> 得到表头和字段关系
 * 自定义样式：使用java8的函数式编程思想，将设定样式的功能通过函数式接口抽离
 * 通过外部编码使样式的设置更加方便
 */
public class ExcelImportExportUtils {

    private static ForkJoinPool pool = new ForkJoinPool(4);

    // 更新线程池
    public static void setPool(ForkJoinPool pool) {
        ExcelImportExportUtils.pool = pool;
    }

    //表头信息
    private enum TableEnum {
        WORK_BOOK, SHEET, TITLE_MAP, TABLE_NAME, TABLE_HEADER, TABLE_HEADER_HEIGHT, RESOURCE, TYPE, TABLE_BODY, PASSWORD, RECORD_ALL_EXCEPTIONS, REVERSE_MAP, TITLE_CACHE
    }

    // 单元格信息
    private enum CellEnum {
        TITLE_NAME, FIELD_NAME, SPLIT_REGEX, REPLACE_ALL, REPLACE_ALL_TYPE, FIELD_TYPE, SIZE, PATTERN, NULLABLE, WIDTH, JOIN, EXCEPTION, COL, ROW, SPLIT, PRIORITY, FORMAT, MAP_ENTRIES, MAP, POSITION_TITLE, OFFSET
    }

    //样式的属性名
    private enum CellStyleEnum {
        FONT_NAME, FONT_SIZE, BG_COLOR, TEXT_ALIGN, LOCKED, HIDDEN, BOLD,
        VERTICAL_ALIGN, WRAP_TEXT, STYLES,
        FORE_COLOR, ROTATION, FILL_PATTERN, AUTO_SHRINK, TOP, BOTTOM, LEFT, RIGHT
    }


    /**
     * 并发导入任务
     */
    private static class ForkJoinImportTask<T> extends RecursiveTask<List<T>> {
        private final int start;
        private final int end;
        private final boolean recordAllException;
        private final Sheet sheet;
        private final int threshold;// 默认1万以后需要拆分
        private final JSONObject fieldInfo;// tableBody
        private final Class<T> clazz; // 泛型
        private static ValidatorFactory vf = Validation.buildDefaultValidatorFactory();
        private static Validator validator = vf.getValidator();


        ForkJoinImportTask(JSONObject fieldInfo, Class<T> clazz, Sheet sheet, int start, int end, boolean recordAllException, int threshold) {
            this.fieldInfo = fieldInfo;
            this.clazz = clazz;
            this.sheet = sheet;
            this.start = start;
            this.end = end;
            this.threshold = threshold;
            this.recordAllException = recordAllException;
        }

        @Override
        protected List<T> compute() {
            int nums = end - start;// 计算有多少行数据

            if (nums <= threshold) {
                // 解析数据并且返回List
                return praseRowsToList();
            } else {
                int middle = (start + end) / 2;

                // 处理start到middle行号内的数据
                ForkJoinImportTask<T> left = new ForkJoinImportTask<>(fieldInfo, clazz, sheet, start, middle, recordAllException, threshold);
                left.fork();
                // 处理middle+1到end行号内的数据
                ForkJoinImportTask<T> right = new ForkJoinImportTask<>(fieldInfo, clazz, sheet, middle + 1, end, recordAllException, threshold);
                right.fork();
                final List<T> leftList = left.join();
                final List<T> rightList = right.join();
                leftList.addAll(rightList);
                return leftList;
            }
        }

        /**
         * 解析从start到end行的数据转换为List
         */
        private List<T> praseRowsToList() throws RuntimeException {
            // 从表头描述信息得到表头的高
            final ArrayList<T> result = new ArrayList<>();

            ArrayList<List<String>> rowOfErrMessage = new ArrayList<>();
            // 按行扫描excel表
            for (int i = start; i <= end; i++) {
                final Row row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }
                JSONObject oneRow = new JSONObject();// 一行数据
                oneRow.put(CellEnum.ROW.name(), i + 1);// 记录行号
                // 错误消息列表
                ArrayList<String> errMessage = new ArrayList<>();

                int countNull = 0;
                // 将Row转换为JSONObject
                int finalI = i;
                fieldInfo.forEach((fieldName, arr) -> {
                    JSONArray fieldDescArr = (JSONArray) arr;
                    String join = null;
                    int replaceType = 0;
                    JSONObject temp = null;
                    // 单个注解对字段的转换逻辑
                    ArrayList<Object> tempList = new ArrayList<>();

                    for (Object obj : fieldDescArr) {
                        JSONObject fieldDesc = (JSONObject) obj;
                        replaceType = fieldDesc.getInteger(CellEnum.REPLACE_ALL_TYPE.name());
                        if (fieldDescArr.size() > 1) {
                            temp = fieldDesc;
                        }
                        try {
                            // 根据注解获取到字符串
                            Object value = getValue(row, finalI, fieldDesc, replaceType);
                            if (value != null) {
                                tempList.add(value);
                                if (join == null) {
                                    join = fieldDesc.getString(CellEnum.JOIN.name());// 进行正则切割
                                } else if (!fieldDesc.getString(CellEnum.JOIN.name()).equals("$$")) {
                                    join = fieldDesc.getString(CellEnum.JOIN.name());
                                }
                            }
                        } catch (Exception e) {
                            errMessage.add(e.getMessage());
                        }
                    }
                    if (tempList.size() == 1) {
                        oneRow.put(fieldName, tempList.get(0));// 添加数据
                    } else if (tempList.size() > 1) {
                        // 多个注解的情况一定是字符串类型的，不然存不下2个不同注解的内容
                        String mergedStr = listJoinString(join, tempList);
                        if (replaceType == 1) {
                            mergedStr = replaceAllOrReplacePart(mergedStr, temp);
                        }
                        oneRow.put(fieldName, mergedStr);
                    }
                });

                // 判断这行数据是否正常
                // 正常情况下count是等于length的，因为每个字段都需要处理
                if (errMessage.size() == 0) {
                    if (countNull != oneRow.size() - 1) {
                        // 进行jsr303校验
                        try {
                            T t = JSONObject.parseObject(oneRow.toJSONString(), clazz);
                            t = CheckLogicUtils.checkNullLogicWithJSR303(t, validator);
                            result.add(t);// 正常情况下添加一条数据
                        } catch (RuntimeException e) {
                            errMessage.add("第" + oneRow.getBigInteger(CellEnum.ROW.name()) + "行数据异常：" + e.getMessage());
                        }
                    }

                } else {
                    // 该行存在error
                    rowOfErrMessage.add(errMessage);
                }
                // 如果没有开启记录所有日志，则该task不会跑完，当达到了100条日志的时候就会抛异常
                if (!recordAllException) {
                    if (rowOfErrMessage.size() > 100) {
                        // 每个线程默认达到100Exception时就结束
                        throw new RuntimeException(Thread.currentThread().getName() + "-->超过100条异常记录\n" + printRowOfException(rowOfErrMessage));
                    }
                }
                // 空行继续扫描,或者正常
            }
            if (!recordAllException) {
                if (rowOfErrMessage.size() == 1) {
                    // 需要终止程序，出现了异常
                    throw new RuntimeException("\n" + listToString(rowOfErrMessage.get(0)));
                } else if (rowOfErrMessage.size() >= 2) {
                    // 需要终止程序，出现了异常
                    throw new RuntimeException("\n\nExcel中有" + rowOfErrMessage.size() + "行数据有Error:\n\n" + printRowOfException(rowOfErrMessage));
                }
            }


            return result;
        }

        private String listJoinString(String join, ArrayList<Object> tempList) {
            if (tempList == null || tempList.size() == 0) {
                return null;
            } else if (tempList.size() == 1) {
                return tempList.get(0).toString();
            } else {
                return tempList.stream().map(String::valueOf).reduce((s1, s2) -> s1 + join + s2).get();
            }
        }


        /**
         * 拼接异常日志
         */
        private String printRowOfException(ArrayList<List<String>> rowOfErrMessage) {
            return rowOfErrMessage.stream().map(ExcelImportExportUtils::listToString).reduce((list1, list2) -> list1 + "\n" + list2).get();
        }


    }


    private static Object getValue(Row row, int i, JSONObject fieldDesc, int replaceType) throws Exception {
        // 得到字段的下标
        Integer index = fieldDesc.getInteger(CellEnum.COL.name());
        String fieldType = fieldDesc.getString(CellEnum.FIELD_TYPE.name());// 字段类型
        Integer width = fieldDesc.getInteger(CellEnum.WIDTH.name());// 得到宽度，如果宽度不为1则需要进行合并多个单元格的内容
        String join = fieldDesc.getString(CellEnum.JOIN.name());// 进行正则切割
        String splitRegex = fieldDesc.getString(CellEnum.SPLIT_REGEX.name());// 进行正则切割
        String value = getValue(row, index, width, join, fieldType);
        if (StringUtils.hasText(splitRegex)) {
            // 判断是否需要正则切割，有则将value进行切割处理
            String[] split = value.split(splitRegex);
            StringBuilder stringBuilder = new StringBuilder();
            for (int j = 0; j < split.length; j++) {
                if (replaceType == 0) {
                    stringBuilder.append(replaceAllOrReplacePart(split[j], fieldDesc));
                }
                if (j + 1 < split.length) {
                    stringBuilder.append(splitRegex);
                }

            }
            // 将处理完后的内容重新赋值给value
            value = stringBuilder.toString();
        } else {
            // 没有内容就直接替换
            if (replaceType == 0) {
                value = replaceAllOrReplacePart(value, fieldDesc);
            }
        }
        //String fieldName = fieldDesc.getString(CellEnum.FIELD_NAME.name());// 字段名称
        String title = fieldDesc.getString(CellEnum.TITLE_NAME.name());// 标题名称
        String exception = fieldDesc.getString(CellEnum.EXCEPTION.name());// 转换异常返回的消息
        String size = fieldDesc.getString(CellEnum.SIZE.name());// 得到规模
        boolean nullable = fieldDesc.getBoolean(CellEnum.NULLABLE.name());
        String pattern = fieldDesc.getString(CellEnum.PATTERN.name());// 获取正则表达式，如果有正则，则进行正则截取value（相当于从单元格中取部分）
        boolean hasText = StringUtils.hasText(value);

        Object castValue = null;
        // 默认字段不可以为空，如果注解过程设置为true则不抛异常
        if (!nullable) {
            // 说明字段不可以为空
            if (!hasText) {
                // 字段不能为空结果为空，这个空字段异常计数+1。除非count==length，然后重新计数，否则就是一行异常数据
                // 进来了，说明不为空字段在excel中为空，所以需要报异常
                throw new Exception("异常：第" + (i + 1) + "行的, " + numToLetter(index + 1) + " 列" + "---单元格不能为空---所在标题：" + title);
            } else {
                try {
                    // 单元格有内容,要么正常、要么异常直接抛不能返回null 直接中止
                    value = patternConvert(pattern, value);
                    castValue = cast(value, fieldType, "异常：第" + (i + 1) + "行的, " + numToLetter(index + 1) + " 列" + exception, size);
                } catch (ClassCastException e) {
                    throw new Exception(e.getMessage() + "\ttype:" + fieldType + "\tvalue:" + value);
                }
            }

        } else {
            // 字段可以为空 （要么正常 要么返回null不会抛异常）
            try {
                // 单元格内容无关紧要。要么正常转换，要么返回null
                value = patternConvert(pattern, value);
                castValue = cast(value, fieldType, null, size);
            } catch (Exception e) {
                //castValue=null;// 本来初始值就是null
            }
        }
        return castValue;
    }

    /**
     * 并发导出任务
     */
    private static class ForkJoinExportTask<T> extends RecursiveTask<Integer> {

        private final int start;
        private final int end;
        private final List<T> subList;
        private final Sheet sheet;
        private final int threshold;// 默认1千以后需要拆分
        private final Function<T, IndexedColors> function;// 功能性函数
        private final JSONObject titleInfo;// 单元格标题列描述信息
        private static ValidatorFactory vf = Validation.buildDefaultValidatorFactory();
        private static Validator validator = vf.getValidator();
        final static ThreadLocal<HashMap<String, CellStyle>> cellStyleThreadLocal = new ThreadLocal<>();

        /**
         * @param subList   数据集合
         * @param titleInfo 单元格标题描述信息
         * @param sheet     excel表
         * @param start     起始行号
         * @param end       结束行号
         * @param threshold 条件因子
         * @param function  功能型函数
         */
        private ForkJoinExportTask(List<T> subList, JSONObject titleInfo, Sheet sheet, int start, int end, int threshold, Function<T, IndexedColors> function) {
            this.subList = subList;
            this.titleInfo = titleInfo;
            this.start = start;
            this.end = end;
            this.threshold = threshold;
            this.sheet = sheet;
            this.function = function;
        }


        @Override
        protected Integer compute() throws RuntimeException {
            int length = end - start;

            if (length <= threshold) {
                return this.subListFilledSheet();
            } else {
                int middle = (start + end) / 2;
                int subIndex = middle - start + 1;// 要包含middle这个位子则需要加1进行截取

                List<T> subList1 = subList.subList(0, subIndex);
                List<T> subList2 = subList.subList(subIndex, subList.size());
                ForkJoinExportTask<T> left = new ForkJoinExportTask<>(subList1, titleInfo, sheet, start, start + subList1.size() - 1, threshold, function);
                left.fork();
                ForkJoinExportTask<T> right = new ForkJoinExportTask<>(subList2, titleInfo, sheet, start + subList1.size(), end, threshold, function);
                right.fork();
                return left.join() + right.join();
            }
        }

        /**
         * 将集合数据填入表格
         */
        private int subListFilledSheet() throws RuntimeException {
            final int length = subList.size();

            short index = IndexedColors.WHITE.getIndex();// 默认白色背景
            // 每一个子列表同一个cellStyle
            CellStyle cellStyle = getCellStyle(index);

            for (int i = 0; i < length; i++) {
                int rowNum = start + i;
                Row row = sheet.getRow(rowNum);
                synchronized (sheet) {
                    if (sheet.getRow(rowNum) == null) {
                        // 创建一行,创建过程涉及到map的线程安全,故需要对sheet加锁
                        row = sheet.createRow(rowNum);
                    }
                }

                AtomicReference<RuntimeException> exception = new AtomicReference<>();
                // 遍历表身体信息
                T t = subList.get(i);
                // 进行jsr303校验数据
                Set<ConstraintViolation<T>> set = validator.validate(t);
                for (ConstraintViolation<T> constraintViolation : set) {
                    throw new RuntimeException("第" + rowNum + "个数据出现异常：" + constraintViolation.getMessage() + ",原数据：" + t);
                }
                if (function != null) {
                    index = function.apply(t).getIndex();
                    cellStyle = getCellStyle(index);
                }

                final JSONObject json = JSONObject.parseObject(JSONObject.toJSONString(t));
                // 可以并行处理单元格
                jsonToOneRow(row, json, exception, cellStyle);

                if (exception.get() != null) {
                    throw exception.get();
                }
            }
            //final Thread thread = Thread.currentThread();
            cellStyleThreadLocal.remove();
            return length;
        }

        /**
         * 获取一个样式，如果没有就创建一个并且进行缓存
         */
        private CellStyle getCellStyle(short index) {
            final HashMap<String, CellStyle> cellStyleMap = cellStyleThreadLocal.get();
            if (cellStyleMap != null) {
                final CellStyle cellStyle2 = cellStyleMap.get(String.valueOf(index));
                if (cellStyle2 != null) {
                    return cellStyle2;// 存在样式直接返回
                }
            } else {
                cellStyleThreadLocal.set(new HashMap<>());
            }
            final Workbook workbook = sheet.getWorkbook();
            final HashMap<String, CellStyle> styleMap = cellStyleThreadLocal.get();
            synchronized (workbook) {
                CellStyle cellStyle = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setFontName("微软雅黑");
                font.setFontHeightInPoints((short) 11);//设置字体大小
                cellStyle.setFont(font);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                cellStyle.setFillForegroundColor(index);// 设置颜色
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                styleMap.put(String.valueOf(index), cellStyle);
            }
            cellStyleThreadLocal.set(styleMap);
            return styleMap.get(String.valueOf(index));

        }

        /**
         * 将数据填入一行
         */
        private void jsonToOneRow(Row row, JSONObject json, AtomicReference<RuntimeException> exception, CellStyle cellStyle) {
            for (Map.Entry<String, Object> entry : titleInfo.entrySet()) {
                final String titleIdx = entry.getKey();
                final Object v = entry.getValue();
                // 标题 索引
                final int titleIndex = Integer.parseInt(titleIdx);
                if (titleIndex < 0) {
                    throw new RuntimeException("导出异常，请校验注解title值和模板标题是否一致，得到的标题索引为" + titleIdx + ".标题的信息是：" + v);
                }
                // 给这个 index单元格 填入 value
                Cell cell = row.getCell(titleIndex);// 得到单元格
                if (cell == null) {
                    cell = row.createCell(titleIndex);
                }
                cell.setCellStyle(cellStyle);
                if (v instanceof JSONArray) {
                    // 多个字段合并成一个单元格内容
                    JSONArray array = (JSONArray) v;
                    String[] linkedFormatString = new String[array.size()];
                    final AtomicInteger atomicInteger = new AtomicInteger(0);
                    final Object[] objects = array.toArray();
                    for (Object obj : objects) {
                        // 处理每一个字段
                        JSONObject fieldDescData = (JSONObject) obj;
                        // 得到转换后的内容
                        final String resultValue = getValueByFieldInfo(json, fieldDescData);
                        linkedFormatString[atomicInteger.getAndIncrement()] = resultValue;// 存入该位置
                    }

                    final StringBuilder stringBuilder = new StringBuilder();
                    for (String s : linkedFormatString) {
                        stringBuilder.append(s);
                    }
                    String value = stringBuilder.toString();// 得到了合并后的内容
                    // 因为合并了多个字段所以肯定是字符串类型的
                    cell.setCellValue(value);

                } else {
                    // 一个字段可能要拆成多个单元格
                    JSONObject jsonObject = (JSONObject) v;
                    final String format = jsonObject.getString(CellEnum.FORMAT.name());
                    final Integer priority = jsonObject.getInteger(CellEnum.PRIORITY.name());
                    final String fieldName = jsonObject.getString(CellEnum.FIELD_NAME.name());
                    final String fieldType = jsonObject.getString(CellEnum.FIELD_TYPE.name());
                    final String size = jsonObject.getString(CellEnum.SIZE.name());
                    final String split = jsonObject.getString(CellEnum.SPLIT.name());
                    Object fieldValue = json.get(fieldName);// 得到这个字段值
                    final Integer width = jsonObject.getInteger(CellEnum.WIDTH.name());

                    if (width > 1) {
                        // 一个字段需要拆分成多个单元格
                        if (StringUtils.hasText(split)) {
                            // 有拆分词,是字符串
                            final String[] splitArray = ((String) fieldValue).split(split);// 先拆分字段
                            if (StringUtils.hasText(format)) {
                                // 有格式化模板
                                final String[] formatStr = format.split(split);// 拆分后的格式化内容
                                for (int j = 0; j < width; j++) {
                                    cell = row.getCell(titleIndex + j);
                                    if (cell == null) {
                                        cell = row.createCell(titleIndex + j);// 得到单元格
                                    }
                                    String formattedStr = formatStr[j].replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellStyle(cellStyle);
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            } else {
                                // 没有格式化模板直接填入内容
                                for (int j = 0; j < width; j++) {
                                    cell = row.getCell(titleIndex + j);
                                    if (cell == null) {
                                        cell = row.createCell(titleIndex + j);// 得到单元格
                                    }
                                    String formattedStr = format.replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellStyle(cellStyle);
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            }

                        } else {
                            // 没有拆分词，本身需要拆分，抛异常
                            exception.set(new RuntimeException(fieldName + "字段的注解上 缺少exportSplit拆分词"));
                        }
                    } else {

                        // 一个字段不需要拆成多个单元格
                        if (StringUtils.hasText(format)) {
                            // 内容存在格式化先进行格式化，然后填入值
                            String replacedStr = format.replace("$" + priority, (String) fieldValue);// 替换字符串
                            cell.setCellValue(replacedStr);// 设置单元格内容
                        } else {
                            // 内容不需要格式化则直接填入(转换一下单位，如果没有就原样返回)
                            setCellValue(cell, fieldValue, fieldType, size);
                        }
                    }
                }
            }

        }

    }


    /**
     * 默认不超过10w行数据使用单线程读，否则将会采用forkjoin并行读
     * 用于读取多个sheet
     * 读取sheet中的数据为List，默认粒度10_0000
     */
    public static <T> List<T> importExcel(Sheet sheet, Class<T> tClass) throws Exception {
        return sheetToList(sheet, tClass, 100000);
    }

    /**
     * 用于读取多个sheet
     * 读取sheet中的数据为List,带任务粒控制因子threshold
     */
    public static <T> List<T> importExcel(Sheet sheet, Class<T> tClass, int threshold) throws RuntimeException {
        return sheetToList(sheet, tClass, threshold);
    }

    /**
     * 超过10w则启用并发导入
     *
     * @param inputStream excel的输入流
     * @param tClass      泛型
     */
    public static <T> List<T> importExcel(InputStream inputStream, Class<T> tClass) throws Exception {
        return importExcel(inputStream, tClass, 100000);
    }

    /**
     * 并发导入
     *
     * @param inputStream 传入的excel输入流
     * @param tClass      泛型
     * @param threshold   并发任务的颗粒度
     */
    public static <T> List<T> importExcel(InputStream inputStream, Class<T> tClass, int threshold) throws Exception {
        ExcelTableHeader tableHeader = tClass.getAnnotation(ExcelTableHeader.class);
        if (tableHeader != null) {
            String tableName = tableHeader.sheetName();
            if (!tableName.trim().equals(ExcelConstants.SHEET1)) {
                //如果不是默认的，则根据名称找到sheet，进行操作
                Sheet sheet = getSheetByInputStream(inputStream, tableName);
                return importExcel(sheet, tClass, threshold);
            }
        }
        // 如果注解中没有信息，默认从第一个sheet读取
        return importExcel(inputStream, 0, tClass, threshold);
    }

    /**
     * 传入待读取的excel文件输入流
     *
     * @param inputStream excel输入流
     * @param sheetIdx    下标
     * @param tClass      泛型
     * @param threshold   并发控制因子
     */
    public static <T> List<T> importExcel(InputStream inputStream, int sheetIdx, Class<T> tClass, int threshold) throws Exception {
        return importExcel(getSheetByInputStream(inputStream, sheetIdx), tClass, threshold);
    }

    /**
     * 导出简单的excel文件
     * <p>
     * 案例地址：https://github.com/1015770492/excel-import-export/blob/master/src/test/java/top/yumbo/test/excel/exportDemo/ExportSimpleExcelDemo.java
     *
     * @param list         数据
     * @param outputStream 输出流
     */
    public static <T> void exportSimpleExcel(List<T> list, TitleBuilders titleBuilders, OutputStream outputStream) throws Exception {
        exportSimpleExcelHighLight(list, titleBuilders, outputStream, null);
    }

    /**
     * @param list          listBean
     * @param titleBuilders 标题信息
     * @param outputStream  导出的文件
     * @param <T>           泛型
     */
    public static <T> void exportSimpleExcelHighLight(List<T> list, TitleBuilders titleBuilders, OutputStream outputStream, Function<T, IndexedColors> function) throws Exception {
        final Workbook workbook = new XSSFWorkbook();
        final Sheet sheet = workbook.createSheet();
        // 生成表头
        generateTableHeader(sheet, titleBuilders);
        // 临时文件
        String tempFile = "temp.xlsx";
        final FileOutputStream fileOutputStream = new FileOutputStream(tempFile);
        // excel写入临时文件
        workbook.write(fileOutputStream);
        // 反过来获取输入流
        final FileInputStream fileInputStream = new FileInputStream(tempFile);
        // 调用导出excel
        listToSheetWithStyle(list, fileInputStream, outputStream, function, 100000);
    }

    /**
     * 生成简单表头
     *
     * @param sheet         表格
     * @param titleBuilders 表头信息
     */
    public static Sheet generateTableHeader(Sheet sheet, TitleBuilders titleBuilders) {
        final Workbook workbook = sheet.getWorkbook();
        final CellStyle cellStyle = workbook.createCellStyle();
        final Font font = workbook.createFont();
        font.setFontName("微软雅黑");
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        final List<List<TitleBuilder>> titleList = titleBuilders.getTitleList();
        // 初始化单元格
        for (int i = 0; i < titleList.size(); i++) {
            Row row = getRow(sheet, i);
            AtomicInteger colLength = new AtomicInteger();
            titleList.get(0).forEach(e -> {
                colLength.addAndGet(e.getWidth());
            });
            for (int j = 0; j < colLength.get(); j++) {
                createCell(row, j);
            }
        }
        // 遍历多少行标题
        for (int i = 0; i < titleList.size(); i++) {
            Row row = getRow(sheet, i);
            // 处理第i行标题
            final List<TitleBuilder> list = titleList.get(i);
            int nextIndex = 0;
            for (int j = 0; j < list.size(); j++) {
                final TitleBuilder titleBuilder = list.get(j);
                // 获取索引和标题，宽度和高度
                int index = titleBuilder.getIndex();
                final String title = titleBuilder.getTitle();
                final int width = titleBuilder.getWidth();
                final int height = titleBuilder.getHeight();
                if (nextIndex <= 0) {
                    // 下一个索引 = 当前索引 也就是0 + 当前宽度
                    nextIndex = index + width;
                } else {
                    // 得到新的当前索引
                    index = nextIndex;
                    titleBuilder.setIndex(index);// 同时更新一下
                    // 下一个索引= 当前索引+当前宽度
                    nextIndex = index + width;

                }
                // 得到当前索引位置的单元格
                Cell cell = row.getCell(index);
                if (cell == null) {
                    cell = row.createCell(index);
                }
                cell.setCellValue(title);
                if (width > 1 || height > 1) {
                    final CellRangeAddress cellAddresses = new CellRangeAddress(i, i + height - 1, index, nextIndex - 1);
                    // 合并单元格
                    sheet.addMergedRegion(cellAddresses);
                }
                cell.setCellStyle(cellStyle);
                // 调整宽度
                final int lastColumnWidth = sheet.getColumnWidth(index);
                final int size = title.getBytes().length;
                if (lastColumnWidth < size * 200) {
                    sheet.setColumnWidth(index, size * 200);
                }
            }
        }

        return sheet;
    }


    /**
     * 通过注解信息得到模板文件
     * 导出excel（默认10000作为粒度，超过10000会使用forkJoin进行拆分任务处理）
     *
     * @param list         待导入的数据集合
     * @param outputStream 导出的文件输出流
     */
    public static <T> void exportExcel(List<T> list, OutputStream outputStream) throws Exception {
        listToSheetWithStyle(list, null, outputStream, null, 10000);
    }


    /**
     * 导出Excel,使用默认样式  （传入list 和 输出流）
     *
     * @param list         待导入的数据集合
     * @param outputStream 导出的文件输出流
     * @param threshold    任务粒度
     */
    public static <T> void exportExcel(List<T> list, OutputStream outputStream, int threshold) throws Exception {
        // 调用的还是高亮行的导出方法
        listToSheetWithStyle(list, null, outputStream, null, threshold);
    }

    /**
     * 传入模板文件的输入流进行导出
     *
     * @param list         数据
     * @param inputStream  excel模板的输入流
     * @param outputStream 返回的输出流
     */
    public static <T> void exportExcel(List<T> list, InputStream inputStream, OutputStream outputStream) throws Exception {
        listToSheetWithStyle(list, inputStream, outputStream, null, 10000);
    }

    /**
     * 并发导出
     *
     * @param list         数据
     * @param inputStream  模板文件输入流
     * @param outputStream 导出后的输出流
     * @param threshold    并发粒度
     */
    public static <T> void exportExcel(List<T> list, InputStream inputStream, OutputStream outputStream, int threshold) throws Exception {
        listToSheetWithStyle(list, inputStream, outputStream, null, threshold);
    }

    /**
     * 多功能的导出，function为null就是默认的导出
     * function不为null就是高亮行的导出
     */
    public static <T> void exportExcelRowHighLight(List<T> list, OutputStream outputStream, Function<T, IndexedColors> function) throws Exception {
        listToSheetWithStyle(list, null, outputStream, function, 10000);
    }

    /**
     * 多功能的导出，function为null就是默认的导出
     * function不为null就是高亮行的导出
     *
     * @param list         数据
     * @param outputStream 导出的文件的输出流
     * @param function     功能性函数，返回颜色值
     */
    public static <T> void exportExcelRowHighLight(List<T> list, OutputStream outputStream, Function<T, IndexedColors> function, int threshold) throws Exception {
        // 调用的还是高亮行导出
        listToSheetWithStyle(list, null, outputStream, function, threshold);
    }


    /**
     * 生成java常量类
     */
    public static void generateTitleBuilders(Sheet sheet, int titleRow) {
        final int lastRowNum = sheet.getLastRowNum();
        List<List<TitleBuilder>> titleList = new ArrayList<>();
        // 先获取所有标题的索引和标题名称（行号根据list下标即可确定）
        for (int i = 0; i < titleRow; i++) {
            final Row row = sheet.getRow(i);
            List<TitleBuilder> list = new ArrayList<>();
            for (int j = 0; j < lastRowNum; j++) {
                final Cell cell = row.getCell(j);
                final String title = getStringCellValue(cell, cell.getCellType());
                if (StringUtils.hasText(title)) {
                    int index = j, height = i;
                    // 找高度
                    for (; height < titleRow; height++) {
                        final String title2 = getStringCellValue(cell, cell.getCellType());
                        if (StringUtils.hasText(title2)) {
                            break;
                        }
                    }
                    // 找宽度
                    for (; index < lastRowNum; index++) {
                        final String title2 = getStringCellValue(cell, cell.getCellType());
                        if (StringUtils.hasText(title2)) {
                            break;
                        }
                    }
//                    list.add(TitleBuilder.builder().index(j).width().height(height).title(title));

                }
            }
            titleList.add(list);
        }
//        sheet.getNumMergedRegions()
    }

    /**
     * 根据输入流返回workbook
     *
     * @param inputStream excel的输入流
     */
    public static Workbook getWorkBookByInputStream(InputStream inputStream) throws Exception {
        if (inputStream == null) {
            throw new NullPointerException("输入流不能为空");
        }
        return WorkbookFactory.create(inputStream);//可以读取xls格式或xlsx格式。
    }


    /**
     * 根据输入流返回sheet
     *
     * @param inputStream 输入流
     * @param idx         第几个sheet
     */
    public static Sheet getSheetByInputStream(InputStream inputStream, int idx) throws Exception {
        final Workbook workbook = getWorkBookByInputStream(inputStream);
        final Sheet sheet = workbook.getSheetAt(idx);
        if (sheet == null) {
            throw new NullPointerException("序号为" + idx + "的sheet不存在");
        }
        return sheet;
    }

    /**
     * 根据名称返回sheet
     *
     * @param inputStream 输入流
     * @param sheetName   sheet的名称
     */
    public static Sheet getSheetByInputStream(InputStream inputStream, String sheetName) throws Exception {
        final Workbook workbook = getWorkBookByInputStream(inputStream);
        final Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new NullPointerException("sheetName为" + sheetName + "的sheet不存在");
        }
        return sheet;
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
        // 根据相同标题填充index
        JSONObject allDescData = filledTitleIndexBySheet(partDescData, sheet);
        // 填充完成后移除 col 为-1的字段，去除多余的判断，提高效率（因为这些字段是不做处理的，没必要循环判断）
        return removeColLessZero(allDescData);
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
        JSONObject titleCache = new JSONObject();

        // 补充table每一项的index信息(需要通过excel来操作)
        tableBodyDesc.forEach((fieldName, cellDescArr) -> {
            // 得到字段的信息
            JSONArray filedInfoArr = (JSONArray) cellDescArr;
            filedInfoArr.forEach((cd) -> {
                JSONObject cellDesc = (JSONObject) cd;
                Integer col = cellDesc.getInteger(CellEnum.COL.name());
                // 扫描包含表头的那几行 记录需要记录的标题所在的索引列，填充INDEX
                for (int i = 0; i < height; i++) {
                    Row row = sheet.getRow(i);// 得到第i行数据（在表头内）
                    if (row == null) {
                        continue;
                    }
                    // 遍历这行所有单元格，然后得到表头进行比较找到标题和注解上的titleName相同的单元格
                    for (Cell cell : row) {
                        // 得到单元格内容（统一为字符串类型）
                        String cellValue = getStringCellValue(cell, String.class.getTypeName());
                        if (col != -1) {
                            titleCache.putIfAbsent(cellValue, col);
                            //说明填了index值，则根据index处理;处理完这个字段了，需要处理下一个
                            break;
                        } else {
                            // 如果标题相同找到了这单元格，获取单元格下标存入
                            if (cellValue.equals(cellDesc.getString(CellEnum.TITLE_NAME.name()))) {
                                int columnIndex = cell.getColumnIndex();// 找到了则取出索引存入jsonObject
                                cellDesc.put(CellEnum.COL.name(), columnIndex); // 补全描述信息
                                titleCache.putIfAbsent(cellValue, columnIndex);
                            }
                        }
                    }
                }
            });

        });
        tableHeaderDesc.put(TableEnum.TITLE_CACHE.name(), titleCache);
        updateTitleIndex(excelDescData);
        return excelDescData;
    }

    /**
     * 重复标题的处理方案：利用不重复标题的位置+当前标题与该标题的偏移
     * 更新部分没有加title 或 index 的字段信息
     * 但是 有些字段加了 position 和 offset信息
     */
    private static void updateTitleIndex(JSONObject excelDescData) {
        JSONObject tableBodyDesc = getExcelBodyDescInfo(excelDescData);
        JSONObject titleCacheInfo = getTitleCacheInfo(excelDescData);
        tableBodyDesc.forEach((fieldName, cda) -> {
            JSONArray fieldInfoArr = (JSONArray) cda;
            fieldInfoArr.forEach(fia -> {
                JSONObject fieldInfo = (JSONObject) fia;
                String positionTitle = fieldInfo.getString(CellEnum.POSITION_TITLE.name());
                if (StringUtils.hasText(positionTitle)) {
                    if (titleCacheInfo.containsKey(positionTitle)) {
                        // 为了不写死 index
                        int newCol = titleCacheInfo.getInteger(positionTitle) + fieldInfo.getInteger(CellEnum.OFFSET.name());
                        fieldInfo.put(CellEnum.COL.name(), newCol);
                    }
                }
            });
        });
    }


    /**
     * 根据MapEntry注解得到字典映射，用于转换
     */
    private static JSONObject getMapByMapEntries(Field field) {
        final top.yumbo.excel.annotation.business.MapEntry[] mapEntries = field.getDeclaredAnnotationsByType(top.yumbo.excel.annotation.business.MapEntry.class);
        final JSONObject mutiMap = new JSONObject();
        final JSONObject map = new JSONObject();
        final JSONObject reverseMap = new JSONObject();

        for (top.yumbo.excel.annotation.business.MapEntry mapEntry : mapEntries) {
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
     * (导入)将sheet解析成List类型的数据
     * （注意这里只是将单元格内容转换为了实体，具体字段可能还不是正确的例如 区域码应该是是具体的编码而不是XX市XX区）
     *
     * @param tClass 传入的泛型
     * @param sheet  表单数据（带表头的）
     * @return 只是将单元格内容转化为List
     */
    private static <T> List<T> sheetToList(Sheet sheet, Class<T> tClass, int threshold) throws RuntimeException {
        sheetNullCheck(sheet);
        sheet.getWorkbook().setActiveSheet(0);
        final JSONObject importInfoByClazz = getImportInfoByClazz(sheet, tClass);
        JSONObject tableHeaderDescInfo = getTableHeaderDescInfo(importInfoByClazz);
        final Integer tableHeight = getTableHeight(tableHeaderDescInfo);
        final JSONObject titleInfo = getExcelBodyDescInfo(importInfoByClazz);

        final int lastRowNum = sheet.getLastRowNum();

        final long start = System.currentTimeMillis();

        if (threshold <= 0) {
            threshold = 10000;
        }
        Boolean recordAllException = tableHeaderDescInfo.getBoolean(TableEnum.RECORD_ALL_EXCEPTIONS.name());
        final ForkJoinImportTask<T> forkJoinAction = new ForkJoinImportTask<>(titleInfo, tClass, sheet, tableHeight, lastRowNum, recordAllException, threshold);
        // 执行任务
        final List<T> result = pool.invoke(forkJoinAction);
        final long end = System.currentTimeMillis();

        System.out.println("forkJoin读转换耗时" + (end - start) + "毫秒");
        return result;
    }

    /**
     * 数字转Excel的字母表示
     */
    private static String numToLetter(int index) {
        if (index < Integer.MAX_VALUE) {
            if (index >= 1 && index <= 26) {
                return intMap.get(index);
            } else {
                int temp = index % 26;
                index -= temp;
                int t = index / 26;
                String s = numToLetter(temp);
                if (t < 26) {
                    return intMap.get(t) + s;
                } else {
                    return numToLetter(t) + s;
                }
            }
        } else {
            throw new OutOfRangeException(index, Integer.highestOneBit(index), Integer.lowestOneBit(index));
        }

    }

    /**
     * excel 字母表示转数字（26进制）
     */
    private static int letterToNum(String index) {
        double num = 0.0;
        index = index.trim();
        try {
            // 如果是数字直接解析返回，如果不是则经过下面转换
            return Integer.parseInt(index);
        } catch (Exception ignored) {
        }

        index = index.toUpperCase();
        for (int i = 0; i < index.length(); i++) {
            int idx = index.length() - i - 1;
            String ch = index.substring(idx, idx + 1);
            if (intMap.containsValue(ch)) {
                num = num + (ch.charAt(0) - 'A' + 1) * Math.pow(26, i);
            } else {
                throw new RuntimeException(index);
            }
        }
        return (int) num - 1;
    }


    private static void createCell(Row row, int j) {
        Cell cell = row.getCell(j);
        if (cell == null) {
            cell = row.createCell(j);
        }
    }

    private static Row getRow(Sheet sheet, int i) {
        Row row = sheet.getRow(i);
        if (row == null) {
            row = sheet.createRow(i);
        }
        return row;
    }

    /**
     * 空校验
     */
    private static void sheetNullCheck(Sheet sheet) {
        if (sheet == null) {
            throw new NullPointerException("sheet不能为Null");
        }
    }

    /**
     * 高亮行的方式导出
     *
     * @param list         数据集合
     * @param inputStream  excel输入流
     * @param outputStream 导出文件的输出流
     * @param function     功能型函数，返回颜色
     * @param threshold    forkJoin的条件因子，拆分任务的粒度
     */
    private static <T> void listToSheetWithStyle(List<T> list, InputStream inputStream, OutputStream outputStream, Function<T, IndexedColors> function, int threshold) throws Exception {
        if (list != null && list.size() > 0 && outputStream != null) {
            Sheet sheet = null;
            Workbook workbook;
            // 一、首先得到excel的模板文件
            // 1、如果存在输入流就从输入流中获取workbook和sheet
            if (inputStream != null) {
                workbook = WorkbookFactory.create(inputStream);//可以读取xls格式或xlsx格式。
                sheet = workbook.getSheetAt(0);
            }
            // 二、得到必要的一些信息用于后续给sheet填充数据
            // 2、传入的sheet如果是null则会从注解中获取模板并且返回必要的信息，如果不为null则会将它存入exportInfo对象中
            final JSONObject exportInfo = getExportInfo(list.get(0).getClass(), sheet);
            final JSONObject titleInfo = getExcelBodyDescInfo(exportInfo);
            // 重新从Json对象中取出workbook和sheet
            workbook = getWorkBook(exportInfo);
            sheet = getSheet(exportInfo);
            // 得到表头的高度
            final Integer height = getTableHeight(getTableHeaderDescInfo(exportInfo));

            final long start = System.currentTimeMillis();
            // 总的数据量
            final int length = list.size();

            if (threshold <= 0) {
                // 默认任务粒度
                threshold = 100000;
            }
            // forkJoin线程池进行导出导出
            final ForkJoinExportTask<T> forkJoinAction = new ForkJoinExportTask<>(list, titleInfo, sheet, height, length + height - 1, threshold, function);
            // 执行任务
            pool.invoke(forkJoinAction);

            final long end = System.currentTimeMillis();

            System.out.println("转换耗时" + (end - start) + "毫秒");
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } else if (list == null || list.size() == 0) {
            throw new NullPointerException("list不能为null 或 list数据量不能为0");
        } else {
            throw new NullPointerException("输出流不能为空");
        }

    }


    private static String listToString(List<String> list) {
        if (list == null || list.size() == 0) {
            return "[]";
        } else if (list.size() == 1) {
            return list.toString();
        } else {
            return "[" + list.stream().reduce((str1, str2) -> str1 + "\n" + str2).get() + "]";
        }
    }

    /**
     * 获取表头的高度
     */
    private static Integer getTableHeight(JSONObject tableHeaderDesc) {
        return tableHeaderDesc.getInteger(TableEnum.TABLE_HEADER_HEIGHT.name());
    }

    /**
     * 得到WorkBook
     */
    private static Workbook getWorkBook(JSONObject fulledDescInfo) {
        return fulledDescInfo.getObject(TableEnum.WORK_BOOK.name(), Workbook.class);
    }


    /**
     * 得到excel表格
     */
    private static Sheet getSheet(JSONObject fulledExcelDescriptionInfo) {
        return fulledExcelDescriptionInfo.getObject(TableEnum.SHEET.name(), Sheet.class);
    }


    /**
     * 得到对象的值(转换的只是格式化后的值)
     *
     * @param entity        实体数据转换的JSONObject
     * @param fieldDescData 字段规则描述数据
     * @return 处理后的字符串
     */
    private static String getValueByFieldInfo(JSONObject entity, JSONObject fieldDescData) {
        final String format = fieldDescData.getString(CellEnum.FORMAT.name());
        final Integer priority = fieldDescData.getInteger(CellEnum.PRIORITY.name());
        final String fieldName = fieldDescData.getString(CellEnum.FIELD_NAME.name());
        // 从对象中得到这个字段值
        String fieldValue = String.valueOf(entity.get(fieldName));// 得到这个字段值

        // 替换字符串
        return format.replace("$" + priority, fieldValue);
    }


    /**
     * 从json中获取Excel身体部分数据
     */
    private static JSONObject getExcelBodyDescInfo(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(TableEnum.TABLE_BODY.name());
    }

    /**
     * 从json中获取到标题头的缓存信息
     */
    private static JSONObject getTitleCacheInfo(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(TableEnum.TABLE_HEADER.name()).getJSONObject(TableEnum.TITLE_CACHE.name());
    }

    /**
     * 从json中获取Excel表头部分数据
     */
    private static JSONObject getTableHeaderDescInfo(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(TableEnum.TABLE_HEADER.name());
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
            tableHeader.put(TableEnum.RECORD_ALL_EXCEPTIONS.name(), excelTableHeaderAnnotation.recordAllExceptions());// 是否开启记录所有异常

            // 2、得到表的Body信息
            for (Field field : fields) {
                ExcelTitleBind[] excelTitleBinds = field.getDeclaredAnnotationsByType(ExcelTitleBind.class);
                if (excelTitleBinds != null) {
                    JSONArray cellDescArr = new JSONArray();
                    for (ExcelTitleBind annotationTitle : excelTitleBinds) {
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
                                cellDesc.put(CellEnum.REPLACE_ALL.name(), annotationTitle.replaceAll());// 是否包含替换所有,默认是替换所有
                                cellDesc.put(CellEnum.REPLACE_ALL_TYPE.name(), annotationTitle.replaceAllType());// 是否包含替换所有,默认是替换所有

                                cellDesc.put(CellEnum.TITLE_NAME.name(), title);// 标题名称
                                cellDesc.put(CellEnum.FIELD_NAME.name(), field.getName());// 字段名称
                                cellDesc.put(CellEnum.FIELD_TYPE.name(), field.getType().getTypeName());// 字段的类型
                                cellDesc.put(CellEnum.COL.name(), letterToNum(annotationTitle.index()));// 默认的索引位置
                                cellDesc.put(CellEnum.WIDTH.name(), annotationTitle.width());// 单元格的宽度（宽度为2代表合并了2格单元格）
                                cellDesc.put(CellEnum.JOIN.name(), annotationTitle.join());// 拼接的字符串 多个单元格以及合并单元格通过它来拼接
                                cellDesc.put(CellEnum.EXCEPTION.name(), annotationTitle.exception());// 校验如果失败返回的异常消息
                                cellDesc.put(CellEnum.SIZE.name(), annotationTitle.size());// 规模,记录规模(亿元/万元)
                                cellDesc.put(CellEnum.PATTERN.name(), annotationTitle.importPattern());// 正则表达式
                                cellDesc.put(CellEnum.NULLABLE.name(), annotationTitle.nullable());// 是否可空
                                cellDesc.put(CellEnum.SPLIT.name(), annotationTitle.exportSplit());// 导出字段的拆分
                                cellDesc.put(CellEnum.FORMAT.name(), annotationTitle.exportFormat());// 导出的模板格式
                                cellDesc.put(CellEnum.PRIORITY.name(), annotationTitle.exportPriority());// 导出拼串的顺序
                                cellDesc.put(CellEnum.POSITION_TITLE.name(), annotationTitle.positionTitle());// 定位
                                cellDesc.put(CellEnum.OFFSET.name(), annotationTitle.offset());// 偏移
                            }
                            cellDescArr.add(cellDesc);
                        }
                    }
                    // 以字段名作为key
                    tableBody.put(field.getName(), cellDescArr);// 存入这个标题名单元格的的描述信息，后面还需要补全INDEX
                }

            }
        }
        excelDescData.put(TableEnum.TABLE_HEADER.name(), tableHeader);// 将表头记录信息注入
        excelDescData.put(TableEnum.TITLE_MAP.name(), titleMap);// 将表头记录信息注入
        excelDescData.put(TableEnum.TABLE_BODY.name(), tableBody);// 将表的body记录信息注入
        return excelDescData;// 返回记录的所有信息
    }

    /**
     * 获取导出所需要的所有信息（标题-字段关系）
     * 注意：会将样式也创建好并且放入返回的jsonObject中，然后直接取
     *
     * @param clazz 传入的泛型，注解信息
     * @param sheet excel表格
     * @return 表格的信息，返回的tableBody中key是每一个标题的索引，value则是由字段信息
     */
    private static JSONObject getExportInfo(Class<?> clazz, Sheet sheet) throws Exception {
        Field[] fields = clazz.getDeclaredFields();// 获取所有字段
        JSONObject tableHeader = new JSONObject();// 表中主体数据信息
        JSONObject tableBody = new JSONObject();// 表中主体数据信息
        JSONObject exportInfo = new JSONObject();// excel的描述数据

        // 1、先得到表头信息
        final ExcelTableHeader tableHeaderAnnotation = clazz.getDeclaredAnnotation(ExcelTableHeader.class);
        if (tableHeaderAnnotation != null) {
            tableHeader.put(TableEnum.TABLE_NAME.name(), tableHeaderAnnotation.sheetName());// 表的名称
            tableHeader.put(TableEnum.TABLE_HEADER_HEIGHT.name(), tableHeaderAnnotation.height());// 表头的高度
            tableHeader.put(TableEnum.RESOURCE.name(), tableHeaderAnnotation.resource());// 模板excel的访问路径
            tableHeader.put(TableEnum.TYPE.name(), tableHeaderAnnotation.type());// xlsx 或 xls 格式
            tableHeader.put(TableEnum.PASSWORD.name(), tableHeaderAnnotation.password());// 模板excel的访问路径
            if (sheet == null) {
                // sheet不存在则从注解信息中获取
                final Workbook workBook = getWorkBookByResource(tableHeaderAnnotation.resource());
                if (workBook != null) {
                    // 根据名称获取sheet，如果名称也没有才获取第一个sheet
                    sheet = workBook.getSheet(tableHeaderAnnotation.sheetName());
                    if (sheet == null) {
                        sheet = workBook.getSheetAt(0);
                        if (sheet == null) {
                            sheet = workBook.createSheet();
                        }
                    }

                }
            }
            sheetNullCheck(sheet);
            exportInfo.put(TableEnum.SHEET.name(), sheet);
            exportInfo.put(TableEnum.WORK_BOOK.name(), sheet.getWorkbook());

            // 2、得到表的Body信息
            for (Field field : fields) {
                final ExcelTitleBind annotationTitle = field.getDeclaredAnnotation(ExcelTitleBind.class);// 获取ExcelCellEnumBindAnnotation注解
                ExcelCellStyle[] annotationStyles = field.getDeclaredAnnotationsByType(ExcelCellStyle.class);// 获取单元格样式注解
                if (annotationTitle != null) {// 找到自定义的注解
                    JSONObject titleDesc = new JSONObject();// 单元格描述信息
                    String title = annotationTitle.title();         // 获取标题，如果标题不存在则不进行处理
                    if (StringUtils.hasText(title)) {
                        // 根据标题找到索引
                        int titleIndexFromSheet = getTitleIndexFromSheet(title, sheet, tableHeaderAnnotation.height());
                        String titleIndexString = String.valueOf(titleIndexFromSheet);// 字符串类型 标题的下标

                        titleDesc.put(CellEnum.TITLE_NAME.name(), annotationTitle.title());// 单元格的宽度（宽度为2代表合并了2格单元格）
                        titleDesc.put(CellEnum.WIDTH.name(), annotationTitle.width());// 单元格的宽度（宽度为2代表合并了2格单元格）
                        titleDesc.put(CellEnum.SPLIT.name(), annotationTitle.exportSplit());// 导出字符串拆分为多个单元格（州市+区县）
                        titleDesc.put(CellEnum.FIELD_NAME.name(), field.getName());// 导出字符串拆分为多个单元格（州市+区县）
                        titleDesc.put(CellEnum.FORMAT.name(), annotationTitle.exportFormat());// 这个字段输出的格式
                        titleDesc.put(CellEnum.PRIORITY.name(), annotationTitle.exportPriority());// 这个字段输出的格式
                        titleDesc.put(CellEnum.SIZE.name(), annotationTitle.size());// 字段规模
                        titleDesc.put(CellEnum.FIELD_TYPE.name(), field.getType().getTypeName());// 字段类型
                        final JSONObject styles = getStylesByAnnotation(annotationStyles, sheet.getWorkbook());// 得到所有样式
                        titleDesc.put(CellStyleEnum.STYLES.name(), styles);// 将样式加入（由多个样式）

                        // 判断是否已经有相同title的注解信息
                        if (tableBody.containsKey(titleIndexString)) {
                            final Object obj = tableBody.get(titleIndexString);// 得到原先的对象（可能是Array也可能是object）
                            if (obj instanceof JSONArray) {
                                JSONArray array = (JSONArray) obj;
                                array.add(titleDesc);// 原先是数组直接添加
                            } else {
                                // 不是数组则转换为数组，并且将原先的对象存入数组
                                final JSONArray array = new JSONArray();
                                array.add(obj);// 存入原先的
                                array.add(titleDesc);// 存入当前的
                                tableBody.put(titleIndexString, array);// 替换为JSONArray
                            }
                        } else {
                            // 没有相同的title则直接添加
                            tableBody.put(titleIndexString, titleDesc);
                        }
                    }
                }
            }
        }
        exportInfo.put(TableEnum.TABLE_HEADER.name(), tableHeader);// 将表头记录信息注入
        exportInfo.put(TableEnum.TABLE_BODY.name(), tableBody);// 将表的body记录信息注入
        return exportInfo;// 返回记录的所有信息
    }

    /**
     * 返回所有样式
     */
    private static JSONObject getStylesByAnnotation(ExcelCellStyle[] annotationStyles, Workbook workbook) {
        final JSONObject styles = new JSONObject();
        for (ExcelCellStyle as : annotationStyles) {
            if (as != null) {
                // 单个样式
                CellStyle cellStyle;
                final CellStyleBuilder styleBuilder = CellStyleBuilder.builder()
                        .fontName(as.fontName())
                        .fontSize(as.fontSize())
                        .textAlign(as.textAlign())
                        .bgColor(as.backgroundColor())
                        .bold(as.bold())
                        .locked(as.locked())
                        .hidden(as.hidden())
                        .wrapText(as.wrapText())
                        .verticalAlignment(as.verticalAlign())
                        .rotation(as.rotation())
                        .fillPatternType(as.fillPatternType())
                        .foregroundColor(as.foregroundColor())
                        .autoShrink(as.autoShrink())
                        .top(as.top())
                        .bottom(as.bottom())
                        .left(as.left())
                        .right(as.right())
                        .build();
                if (workbook == null) {
                    cellStyle = styleBuilder.getCellStyle();
                } else {
                    cellStyle = styleBuilder.getCellStyle(workbook);
                }
                // 将这个样式存入，下次通过id取出
                styles.put(as.id(), cellStyle);
            }
        }
        return styles;
    }

    /**
     * 获取excel工作簿
     *
     * @param resourcePath 资源路径
     */
    private static Workbook getWorkBookByResource(String resourcePath) throws Exception {
        String protoPattern = "(.*)://.*";// 得到协议名称
        final Pattern httpPattern = Pattern.compile(protoPattern);
        final Matcher matcher = httpPattern.matcher(resourcePath);
        if (matcher.find()) {
            final String proto = matcher.group(matcher.groupCount());// 得到协议名
            if (StringUtils.hasText(proto)) {
                InputStream inputStream;
                if ("http".equals(proto) || "https".equals(proto)) {
                    RestTemplate restTemplate = new RestTemplate();
                    HttpHeaders headers = new HttpHeaders();//创建请求头对象
                    HttpEntity<String> entity = new HttpEntity<>("", headers);
                    ResponseEntity<byte[]> obj = restTemplate.exchange(resourcePath, HttpMethod.GET, entity, byte[].class);
                    final byte[] body = obj.getBody();
                    //获取请求中的输入流
                    //java9新特性try升级 自动关闭流
                    if (body == null) {
                        return null;
                    }
                    try (final InputStream is = new ByteArrayInputStream(body)) {
                        inputStream = is;
                    }
                } else {
                    final String[] split = resourcePath.split("://");
                    if (split[1].startsWith("/")) {
                        // 绝对路径
                        inputStream = new FileInputStream(split[1]);
                    } else {
                        // 是相对路径，springboot环境下，打成jar也有效
                        ClassPathResource classPathResource = new ClassPathResource(split[1]);
                        inputStream = classPathResource.getInputStream();
                    }

                }
                return getWorkBookByInputStream(inputStream);
            }
            throw new IllegalArgumentException("请带上协议头例如http:// 或者 https://");
        } else {
            throw new IllegalArgumentException("资源地址不正确，配置的资源地址：" + resourcePath);
        }
    }


    /**
     * 根据标题获取标题在excel中的下标位子
     *
     * @param title      标题名
     * @param sheet      excel的表格
     * @param scanRowNum 表头末尾所在的行数（扫描0-scanRowNum所有行）
     * @return 标题的索引
     */
    private static int getTitleIndexFromSheet(String title, Sheet sheet, int scanRowNum) {
        int index = -1;
        // 扫描包含表头的那几行 记录需要记录的标题所在的索引列，填充INDEX
        for (int i = 0; i < scanRowNum; i++) {
            Row row = sheet.getRow(i);
            // 得到第i行数据（在表头内）
            // 遍历这行所有单元格，然后得到表头进行比较找到标题和注解上的titleName相同的单元格
            for (Cell cell : row) {
                // 得到单元格内容（统一为字符串类型）
                String titleName = getStringCellValue(cell, String.class.getTypeName());
                // 如果标题相同找到了这单元格，获取单元格下标存入
                if (title.equals(titleName)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return index;
    }

    private static String getStringCellValue(Cell cell, CellType cellType) {
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
     * @param row       被取出的行
     * @param fieldInfo index 索引位置 width 索引位置+width确定取那几列
     * @param join      拼接的字符串
     * @param type
     * @return
     */
    private static String getMergedString(Row row, List<List<Integer>> fieldInfo, String join, String type) {
        if (fieldInfo == null || fieldInfo.size() == 0) {
            return "";
        } else if (fieldInfo.size() == 1) {
            return getValue(row, fieldInfo.get(0).get(0), fieldInfo.get(0).get(1), join, type);
        } else {
            return fieldInfo.stream()
                    .map(e -> getValue(row, e.get(0), e.get(1), join, type))
                    .reduce((s1, s2) -> s1 + join + join + s2).get();
        }
    }

    /**
     * 获取单元格内容（逗号分隔）
     *
     * @param row   被取出的行
     * @param index 索引位置
     * @param join  拼接的字符串
     * @param width 索引位置+width确定取那几列
     * @return 返回合并单元格的内容（单个的则传width=1即可）
     */
    private static String getValue(Row row, Integer index, Integer width, String join, String type) {
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
                    cellValue.append(join);
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
     * 移除Col下标为-1的字段描述信息。这样可以少一些循环的判断。
     */
    private static JSONObject removeColLessZero(JSONObject allDescData) {
        JSONObject tableBodyDesc = getExcelBodyDescInfo(allDescData);
        JSONObject newBodyDesc = new JSONObject();
        // 补充table每一项的index信息
        tableBodyDesc.forEach((fieldName, cda) -> {
            JSONArray cellDescArr = (JSONArray) cda;
            JSONArray newArr = new JSONArray();
            cellDescArr.forEach(cellDesc -> {
                if (((JSONObject) cellDesc).getInteger(CellEnum.COL.name()) >= 0) {
                    newArr.add(cellDesc);
                }
            });
            if (newArr.size() > 1) {
                updateFieldArr(newArr);
            }
            newBodyDesc.put(fieldName, newArr);
        });
        allDescData.put(TableEnum.TABLE_BODY.name(), newBodyDesc);
        return allDescData;
    }

    private static void updateFieldArr(JSONArray newArr) {
        HashSet<Integer> replaceSet = new HashSet<>();
        int replaceType = 0;
        for (int i = 0; i < newArr.size(); i++) {
            JSONObject fieldDesc = newArr.getJSONObject(i);
            int[] replaceAll = fieldDesc.getObject(CellEnum.REPLACE_ALL.name(), int[].class);// 替换所有还是替换部分
            replaceType += fieldDesc.getInteger(CellEnum.REPLACE_ALL_TYPE.name());// 替换所有还是替换部分
            for (int i1 : replaceAll) {
                replaceSet.add(i1);
            }
        }
        replaceType = (replaceType == 0) ? 0 : 1;
        for (int i = 0; i < newArr.size(); i++) {
            JSONObject update = newArr.getJSONObject(i);
            update.put(CellEnum.REPLACE_ALL_TYPE.name(), replaceType);
            if (replaceType == 1) {
                // 1.整体替换，合并REPLACE_ALL
                update.put(CellEnum.REPLACE_ALL.name(),
                        replaceSet.stream().mapToInt(Integer::valueOf).filter(x -> x > 0).toArray());
            }//2.只影响到自己
        }
    }

    @Data
    @Builder
    private static class MapEntry implements Comparable<MapEntry> {
        private String key;
        private String value;
        private int id;
        private int idx;
        private int length;

        @Override
        public int compareTo(MapEntry o) {
            return this.getIdx() - o.getIdx();
        }
    }

    /**
     * 得到映射结果
     * 大都情况下containsReplaceAll=true
     * 如果containsReplaceAll=false，需要注意替换部分，必须MapEntry中本身的key不能包含，否则就会替换错误的字典项
     */
    private static String replaceAllOrReplacePart(String value, JSONObject fieldDesc) {
        if (fieldDesc == null) {
            return null;
        }
        JSONObject map = fieldDesc.getJSONObject(CellEnum.MAP.name());
        int[] replaceAll = fieldDesc.getObject(CellEnum.REPLACE_ALL.name(), int[].class);
        if (map == null || !StringUtils.hasText(value)) {
            return value.trim();
        }
        // 转换为字典项
        value = value.trim();// 去掉首尾多余空格等无实意符合
        if (replaceAll.length == 0) {
            // 是完全替换

            for (Map.Entry<String, Object> mapEntry : map.entrySet()) {
                if (value.contains(mapEntry.getKey())) {
                    value = value.replaceAll(mapEntry.getKey(), mapEntry.getValue().toString());
                }
            }

        } else {
            StringBuilder valSb = new StringBuilder(value);
            List<MapEntry> mapEntryList = new ArrayList<>();
            // 先找出有多少个
            for (Map.Entry<String, Object> mapEntry : map.entrySet()) {
                int fromIdx = 0;
                while ((fromIdx < valSb.length()) && valSb.indexOf(mapEntry.getKey(), fromIdx) != -1) {
                    int idx = valSb.indexOf(mapEntry.getKey(), fromIdx);
                    int length = mapEntry.getKey().length();
                    mapEntryList.add(MapEntry.builder().idx(idx).length(length)
                            .key(mapEntry.getKey()).value(mapEntry.getValue().toString()).build());
                    fromIdx = idx + length;
                }
            }
            AtomicInteger id = new AtomicInteger(0);
            // 排完序后进行编号
            List<MapEntry> sortedMapEntry = mapEntryList.stream().sorted().peek(e -> e.setId(id.getAndIncrement())).collect(Collectors.toList());
            // 得到配置需要替换的
            int[] sortedReplaceAll = Arrays.stream(replaceAll).sorted().toArray();

            for (int i = sortedReplaceAll.length - 1; i >= 0; i--) {
                // 先替换最后一个,这样就不会影响前面的
                if (sortedReplaceAll[i] <= sortedMapEntry.size()) {
                    MapEntry mapEntry = sortedMapEntry.get(sortedReplaceAll[i] - 1);
                    valSb.replace(mapEntry.getIdx(), mapEntry.getIdx() + mapEntry.getLength(), mapEntry.getValue());
                }
            }
            value = valSb.toString();
        }
        return value;
    }

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

    /**
     * 得到类型处理后的值字符串
     *
     * @param inputValue 输入的值
     * @param type       应该的类型
     * @param size       规模
     */
    private static void setCellValue(Cell cell, Object inputValue, String type, String size) {

        final Class<?> fieldType = clazzMap.get(type);
        if (fieldType == BigDecimal.class || fieldType == Double.class || fieldType == Float.class
                || fieldType == Long.class || fieldType == Integer.class || fieldType == Short.class) {

            // 数值类型的统一用double（因为cell的api只提供了double）
            BigDecimal bigDecimal = new BigDecimal(String.valueOf(inputValue));
            final BigDecimal resultBigDecimal = BigDecimalUtils.bigDecimalDivBigDecimalFormatTwo(bigDecimal, new BigDecimal(size));
            cell.setCellValue(resultBigDecimal.doubleValue());

        } else if (Date.class == fieldType) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                cell.setCellValue(sdf.parse(inputValue.toString()));
            } catch (ParseException p) {
                throw new RuntimeException("Date类型的字段请加上 @JSONField(format=\"yyyy-MM-dd HH:mm:ss\")");
            }
        } else if (LocalDate.class == fieldType) {

            cell.setCellValue(LocalDate.parse(String.valueOf(inputValue)));
        } else if (LocalDateTime.class == fieldType) {
            cell.setCellValue(LocalDateTime.parse(String.valueOf(inputValue)));
        } else if (GregorianCalendar.class == fieldType) {
            cell.setCellValue((GregorianCalendar) inputValue);
        } else {
            // 字符串或字符类型
            cell.setCellValue((String) inputValue);
        }
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
                obj = new BigDecimal(value).multiply(new BigDecimal(size)).stripTrailingZeros();// 乘以规模
            } else if (aClass == String.class) {
                if (!"1".equals(size)) {
                    // 如果size有值，则进行转换
                    obj = new BigDecimal(value).multiply(new BigDecimal(size)).stripTrailingZeros().doubleValue();
                } else {
                    obj = value;//直接返回字符串
                }
            } else if (aClass == Integer.class) {
                obj = new BigDecimal(value).intValue();
            } else if (aClass == Long.class) {
                obj = Long.parseLong(value.split("\\.")[0]);// 小数点以后的不要
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
            throw new ClassCastException("类型转换异常: value = " + inputValue + ",字段type = " + aClass.getTypeName() + ",size=" + size + " 位置：" + exception);
        }

        return obj;
    }

}
