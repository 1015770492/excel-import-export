package top.yumbo.excel.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.StringUtils;
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelCellStyle;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.entity.CellStyleEntity;
import top.yumbo.excel.interfaces.SerializableFunction;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author jinhua
 * @date 2021/5/21 21:51
 */
public class ExcelImportExportUtils {

    /**
     * 表头信息
     */
    public enum TableEnum {
        TABLE_NAME, TABLE_HEADER, TABLE_HEADER_HEIGHT, RESOURCE, TABLE_BODY, PASSWORD;
    }

    /**
     * 单元格信息
     */
    public enum ExcelCellEnum {
        TITLE_NAME, FIELD_NAME, FIELD_TYPE, SIZE, PATTERN, NULLABLE, WIDTH, EXCEPTION, COL, ROW, SPLIT, PRIORITY, FORMAT;
    }

    /**
     * 样式的属性名
     */
    public enum CellStyleEnum {
        FONT_NAME, FONT_SIZE, BG_COLOR, TEXT_ALIGN, LOCKED, HIDDEN,
        VERTICAL_ALIGN, WRAP_TEXT,
        FORE_COLOR, ROTATION, FILL_PATTERN, AUTO_SHRINK, TOP, BOTTOM, LEFT, RIGHT;
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
        JSONObject fulledExcelDescData = getFulledExcelDescriptionInfo(tClass, sheet);
        // 根据所有已知信息将excel转换为JsonArray数据
        return sheetToJSONArray(fulledExcelDescData, sheet);
    }

    /**
     * 填充List数据数据
     *
     * @param list  数据集
     * @param sheet 待填入的excel表格
     * @throws Exception 抛出的异常
     */
    public static <T> void filledListToSheet(List<T> list, Sheet sheet) throws Exception {
        if (list != null && sheet != null && list.size() > 0) {
            final JSONArray jsonArray = listToJSONArray(list);
            final JSONObject descriptionByClazzAndSheet = getExportDescriptionByClazzAndSheet(list.get(0).getClass(), sheet);
            System.out.println(descriptionByClazzAndSheet);
            filledJSONArrayToSheet(jsonArray, descriptionByClazzAndSheet, sheet);
        } else if (list == null) {
            throw new Exception("list不能为空");
        } else {
            throw new Exception("sheet不能为空");
        }
    }

    public static <T> void filledListToSheetWithCellStyle(List<T> list, Sheet sheet) throws Exception {
        filledListToSheetWithCellStyleByFieldPredicate(list, null, null, x -> false, sheet);
    }

    public static <T> void filledListToSheetWithCellStyleByFunction(List<T> list, List<CellStyle> cellStyleList, Function<T, Integer> function, Sheet sheet) throws Exception {
        JSONObject fulledExcelDescriptionInfo = getExportDescriptionByClazzAndSheet(list.get(0).getClass(), sheet);// 包含了索引填充

        JSONObject tableHeaderDesc = getExcelHeaderDescInfo(fulledExcelDescriptionInfo);// 表头描述信息
        JSONObject tableBodyDesc = getExcelBodyDescInfo(fulledExcelDescriptionInfo);// 表身体描述信息

        final Integer height = tableHeaderDesc.getInteger(TableEnum.TABLE_HEADER_HEIGHT.name());// 得到表头占多少行
        final JSONArray jsonArray = listToJSONArray(list);// list转jsonArray
        // 一行一行填充
        for (int i = 0; i < jsonArray.size(); i++) {
            int rowNum = height + i;
            final Row[] row = {sheet.getRow(rowNum)};// 创建一行数据
            final JSONObject json = (JSONObject) jsonArray.get(i);// 得到这条数据
            AtomicReference<Exception> exception = new AtomicReference<>();
            int finalI = i;
            tableBodyDesc.forEach((index, v) -> {
                if (row[0] == null) {
                    row[0] = sheet.createRow(rowNum);
                }
                // 得到单元格后面给这个 index单元格 填入 value
                Cell cell = row[0].getCell(Integer.parseInt(index));// 得到单元格
                if (cell == null) {
                    cell = row[0].createCell(Integer.parseInt(index));
                }
                Integer styleIndex = function.apply(list.get(finalI));// 根据业务进行处理返回哪一个样式
                cell.setCellStyle(cellStyleList.get(styleIndex));// 使用自定义样式代替
                if (v instanceof JSONArray) {
                    // 多个字段合并成一个单元格内容
                    JSONArray array = (JSONArray) v;
                    String[] linkedFormatString = new String[array.size()];
                    final AtomicInteger atomicInteger = new AtomicInteger(0);
                    array.forEach(obj -> {
                        // 处理每一个字段
                        JSONObject fieldDescData = (JSONObject) obj;
                        // 得到转换后的内容
                        final String resultValue = getFieldValue(json, fieldDescData);
                        linkedFormatString[atomicInteger.getAndIncrement()] = resultValue;// 存入该位置
                    });
                    final StringBuilder stringBuilder = new StringBuilder();
                    for (String s : linkedFormatString) {
                        stringBuilder.append(s);
                    }
                    String value = stringBuilder.toString();// 得到了合并后的内容
                    cell.setCellValue(value);

                } else {
                    // 一个字段可能要拆成多个单元格
                    JSONObject jsonObject = (JSONObject) v;
                    final String format = jsonObject.getString(ExcelCellEnum.FORMAT.name());
                    final Integer priority = jsonObject.getInteger(ExcelCellEnum.PRIORITY.name());
                    final String fieldName = jsonObject.getString(ExcelCellEnum.FIELD_NAME.name());
                    final String fieldType = jsonObject.getString(ExcelCellEnum.FIELD_TYPE.name());
                    final String size = jsonObject.getString(ExcelCellEnum.SIZE.name());
                    final String split = jsonObject.getString(ExcelCellEnum.SPLIT.name());

                    String fieldValue = String.valueOf(json.get(fieldName));// 得到这个字段值
                    final Integer width = jsonObject.getInteger(ExcelCellEnum.WIDTH.name());

                    if (width > 1) {
                        // 一个字段需要拆分成多个单元格
                        if (StringUtils.hasText(split)) {
                            // 有拆分词
                            final String[] splitArray = fieldValue.split(split);// 先拆分字段
                            if (StringUtils.hasText(format)) {
                                // 有格式化模板
                                final String[] formatStr = format.split(split);// 拆分后的格式化内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(Integer.parseInt(index) + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(Integer.parseInt(index) + j);// 得到单元格
                                    }
                                    cell.setCellStyle(cellStyleList.get(styleIndex));// 使用自定义样式代替
                                    String formattedStr = formatStr[j].replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            } else {
                                // 没有格式化模板直接填入内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(Integer.parseInt(index) + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(Integer.parseInt(index) + j);// 得到单元格
                                    }
                                    cell.setCellStyle(cellStyleList.get(styleIndex));// 使用自定义样式代替
                                    String formattedStr = format.replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            }

                        } else {
                            // 没有拆分词，本身需要拆分，抛异常
                            exception.set(new Exception(fieldName + "字段的注解上 缺少exportSplit拆分词"));
                        }
                    } else {
                        // 一个字段不需要拆成多个单元格
                        if (StringUtils.hasText(format)) {
                            // 内容存在格式化先进行格式化，然后填入值
                            String replacedStr = format.replace("$" + priority, fieldValue);// 替换字符串
                            cell.setCellValue(replacedStr);// 设置单元格内容
                        } else {
                            // 内容不需要格式化则直接填入(转换一下单位，如果没有就原样返回)
                            final String result = castForExport(fieldValue, fieldType, size);
                            cell.setCellValue(result);
                        }
                    }
                }
            });

            if (exception.get() != null) {
                throw exception.get();
            }
        }
    }

    /**
     * 生成自定义的表格
     *
     * @param list      数据集合
     * @param cellStyle 自定义的样式
     * @param predicate 断言器
     * @param sheet     待填入的表
     */
    public static <T, R> void filledListToSheetWithCellStyleByFieldPredicate(List<T> list, CellStyle cellStyle, SerializableFunction<T, R> function, Predicate<T> predicate, Sheet sheet) throws Exception {
        JSONObject fulledExcelDescriptionInfo = getExportDescriptionByClazzAndSheet(list.get(0).getClass(), sheet);// 包含了索引填充
        JSONObject tableHeaderDesc = getExcelHeaderDescInfo(fulledExcelDescriptionInfo);// 表头描述信息
        JSONObject tableBodyDesc = getExcelBodyDescInfo(fulledExcelDescriptionInfo);// 表身体描述信息

        final Field field = ReflectionUtil.getField(function);
        final ExcelCellBind annotation = field.getAnnotation(ExcelCellBind.class);

        final Integer height = tableHeaderDesc.getInteger(TableEnum.TABLE_HEADER_HEIGHT.name());// 得到表头占多少行
        final int titleIndex = getTitleIndexFromSheet(annotation.title(), sheet, height);// 得到这个字段的表格索引，这个索引列需要进行断言
        final JSONArray jsonArray = listToJSONArray(list);// list转jsonArray
        // 一行一行填充
        for (int i = 0; i < jsonArray.size(); i++) {
            int rowNum = height + i;
            final Row[] row = {sheet.getRow(rowNum)};// 获取一行一行数据
            if (row[0] == null) {
                row[0] = sheet.createRow(rowNum);
            }
            final JSONObject json = (JSONObject) jsonArray.get(i);// 得到这条数据
            AtomicReference<Exception> exception = new AtomicReference<>();
            int finalI = i;

            tableBodyDesc.forEach((index, v) -> {

                // 得到单元格后面给这个 index单元格 填入 value
                Cell cell = row[0].getCell(Integer.parseInt(index));// 得到单元格
                if (cell == null) {
                    cell = row[0].createCell(Integer.parseInt(index));
                }
                final boolean test = predicate.test(list.get(finalI));// 是否启用样式

                if (v instanceof JSONArray) {
                    // 一个标题，多个字段。 多个字段内容合并成一个单元格内容
                    JSONArray array = (JSONArray) v;
                    String[] linkedFormatString = new String[array.size()];
                    final AtomicInteger atomicInteger = new AtomicInteger(0);
                    array.forEach(obj -> {
                        // 处理每一个字段
                        JSONObject fieldDescData = (JSONObject) obj;
                        // 得到转换后的内容
                        final String resultValue = getFieldValue(json, fieldDescData);
                        linkedFormatString[atomicInteger.getAndIncrement()] = resultValue;// 存入该位置
                    });
                    final StringBuilder stringBuilder = new StringBuilder();
                    for (String s : linkedFormatString) {
                        stringBuilder.append(s);
                    }
                    String value = stringBuilder.toString();// 得到了合并后的内容
                    if (test && Integer.parseInt(index) == titleIndex) {
                        cell.setCellStyle(cellStyle);// 使用自定义样式代替
                    } else {
                        JSONObject cellStyleDesc = array.getJSONObject(0);
                        // 使用注解上的默认样式，如果没有则给一个默认样式
                        setCellStyle(cell, cellStyleDesc);
                    }
                    cell.setCellValue(value);

                } else {
                    // 一个标题一个字段，一个字段拆分多个单元格
                    JSONObject jsonObject = (JSONObject) v;
                    final String format = jsonObject.getString(ExcelCellEnum.FORMAT.name());
                    final Integer priority = jsonObject.getInteger(ExcelCellEnum.PRIORITY.name());
                    final String fieldName = jsonObject.getString(ExcelCellEnum.FIELD_NAME.name());
                    final String fieldType = jsonObject.getString(ExcelCellEnum.FIELD_TYPE.name());
                    final String size = jsonObject.getString(ExcelCellEnum.SIZE.name());
                    final String split = jsonObject.getString(ExcelCellEnum.SPLIT.name());

                    String fieldValue = String.valueOf(json.get(fieldName));// 得到这个字段值
                    final Integer width = jsonObject.getInteger(ExcelCellEnum.WIDTH.name());

                    if (width > 1) {
                        // 一个字段需要拆分成多个单元格
                        if (StringUtils.hasText(split)) {
                            // 有拆分词
                            final String[] splitArray = fieldValue.split(split);// 先拆分字段
                            if (StringUtils.hasText(format)) {
                                // 有格式化模板
                                final String[] formatStr = format.split(split);// 拆分后的格式化内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(Integer.parseInt(index) + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(Integer.parseInt(index) + j);// 得到单元格
                                    }
                                    if (test && Integer.parseInt(index) == titleIndex) {
                                        cell.setCellStyle(cellStyle);// 使用自定义样式代替
                                    } else {
                                        JSONObject cellStyleDesc = (JSONObject) v;
                                        // 使用注解上的默认样式，如果没有则给一个默认样式
                                        setCellStyle(cell, cellStyleDesc);
                                    }
                                    String formattedStr = formatStr[j].replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            } else {
                                // 没有格式化模板直接填入内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(Integer.parseInt(index) + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(Integer.parseInt(index) + j);// 得到单元格
                                    }
                                    if (test && Integer.parseInt(index) == titleIndex) {
                                        cell.setCellStyle(cellStyle);// 使用自定义样式代替
                                    } else {
                                        JSONObject cellStyleDesc = (JSONObject) v;
                                        // 使用注解上的默认样式，如果没有则给一个默认样式
                                        setCellStyle(cell, cellStyleDesc);
                                    }
                                    String formattedStr = format.replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            }

                        } else {
                            // 没有拆分词，本身需要拆分，抛异常
                            exception.set(new Exception(fieldName + "字段的注解上 缺少exportSplit拆分词"));
                        }
                    } else {
                        if (test && Integer.parseInt(index) == titleIndex) {
                            cell.setCellStyle(cellStyle);// 使用自定义样式代替
                        } else {
                            JSONObject cellStyleDesc = (JSONObject) v;
                            // 使用注解上的默认样式，如果没有则给一个默认样式
                            setCellStyle(cell, cellStyleDesc);
                        }
                        // 一个字段不需要拆成多个单元格
                        if (StringUtils.hasText(format)) {
                            // 内容存在格式化先进行格式化，然后填入值
                            String replacedStr = format.replace("$" + priority, fieldValue);// 替换字符串
                            cell.setCellValue(replacedStr);// 设置单元格内容
                        } else {
                            // 内容不需要格式化则直接填入(转换一下单位，如果没有就原样返回)
                            final String result = castForExport(fieldValue, fieldType, size);
                            cell.setCellValue(result);
                        }
                    }
                }
            });

            if (exception.get() != null) {
                throw exception.get();
            }
        }
    }

    /**
     * 设置单元格样式
     */
    private static void setCellStyle(Cell cell, JSONObject cellStyle) {
        final String fontName = cellStyle.getString(CellStyleEnum.FONT_NAME.name());
        final Short fontSize = cellStyle.getShort(CellStyleEnum.FONT_SIZE.name());
        final Short bgColor = cellStyle.getShort(CellStyleEnum.BG_COLOR.name());
        final Short foreColor = cellStyle.getShort(CellStyleEnum.FORE_COLOR.name());
        final Short rotation = cellStyle.getShort(CellStyleEnum.ROTATION.name());
        final boolean locked = cellStyle.getBooleanValue(CellStyleEnum.LOCKED.name());
        final boolean wrapText = cellStyle.getBoolean(CellStyleEnum.WRAP_TEXT.name());
        final boolean hidden = cellStyle.getBoolean(CellStyleEnum.HIDDEN.name());
        final boolean shrink = cellStyle.getBoolean(CellStyleEnum.AUTO_SHRINK.name());

        final HorizontalAlignment textAlign = cellStyle.getObject(CellStyleEnum.TEXT_ALIGN.name(), HorizontalAlignment.class);
        final VerticalAlignment verticalAlignment = cellStyle.getObject(CellStyleEnum.VERTICAL_ALIGN.name(), VerticalAlignment.class);
        final FillPatternType fillPatternType = cellStyle.getObject(CellStyleEnum.FILL_PATTERN.name(), FillPatternType.class);
        final BorderStyle top = cellStyle.getObject(CellStyleEnum.TOP.name(), BorderStyle.class);
        final BorderStyle bottom = cellStyle.getObject(CellStyleEnum.BOTTOM.name(), BorderStyle.class);
        final BorderStyle left = cellStyle.getObject(CellStyleEnum.LEFT.name(), BorderStyle.class);
        final BorderStyle right = cellStyle.getObject(CellStyleEnum.RIGHT.name(), BorderStyle.class);
        setCellStyle(cell, fontName, fontSize, locked, hidden, textAlign, bgColor, foreColor, rotation, verticalAlignment, fillPatternType, top, bottom, left, right, wrapText, shrink);
    }

    private static void setCellStyle(Cell cell,
                                     String fontName, Short fontSize, Boolean locked, Boolean hidden,
                                     HorizontalAlignment textAlign,
                                     Short bgColor, Short foreignColor,
                                     Short rotation, VerticalAlignment verticalAlignment, FillPatternType fillPatternType,
                                     BorderStyle top, BorderStyle bottom, BorderStyle left, BorderStyle right, Boolean wrapText,
                                     Boolean autoShrink
    ) {

        CellStyle cellStyle = getCellStyle(cell.getSheet().getWorkbook(), fontName, (int) fontSize, locked, hidden, textAlign, (int) bgColor, (int) foreignColor, (int) rotation, verticalAlignment, fillPatternType, top, bottom, left, right, wrapText, autoShrink);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 创建一个单元格样式
     *
     * @param fontName          字体
     * @param fontSize          字体大小
     * @param locked            是否可编辑
     * @param hidden            是否隐藏样式
     * @param textAlign         水平对齐方式（居中）
     * @param bgColor           背景颜色
     * @param foreignColor      前景颜色
     * @param rotation          文字旋转角度
     * @param verticalAlignment 垂直对齐方式
     * @param fillPatternType   填充图案
     * @param top               上边框
     * @param bottom            下边框
     * @param left              左边框
     * @param right             右边框
     * @param wrapText          是否自动换行
     * @param autoShrink        是否自动调整大小
     */
    public static CellStyle getCellStyle(Workbook workbook, String fontName, Integer fontSize, Boolean locked, Boolean hidden, HorizontalAlignment textAlign,
                                         Integer bgColor, Integer foreignColor,
                                         Integer rotation, VerticalAlignment verticalAlignment, FillPatternType fillPatternType,
                                         BorderStyle top, BorderStyle bottom, BorderStyle left, BorderStyle right, Boolean wrapText,
                                         Boolean autoShrink
    ) {
        return CellStyleEntity.builder().fontName(fontName).fontSize(fontSize).locked(locked)
                .hidden(hidden).textAlign(textAlign).bgColor(bgColor).foregroundColor(foreignColor)
                .rotation(rotation).verticalAlignment(verticalAlignment)
                .fillPatternType(fillPatternType)
                .top(top).bottom(bottom).left(left).right(right).wrapText(wrapText)
                .autoShrink(autoShrink).build().getCellStyle(workbook);
    }

    /**
     * list转JSONArray
     *
     * @param list 集合
     */
    private static JSONArray listToJSONArray(List<?> list) {
        return JSONObject.parseArray(JSONObject.toJSONString(list));
    }


    /**
     * 存在模板的情况下
     * 将数据填充进入Excel表格
     */
    public static void filledJSONArrayToSheet(JSONArray jsonArray, JSONObject excelDescData, Sheet sheet) throws Exception {
        JSONObject tableHeaderDesc = getExcelHeaderDescInfo(excelDescData);
        JSONObject tableBodyDesc = getExcelBodyDescInfo(excelDescData);
        final Integer height = tableHeaderDesc.getInteger(TableEnum.TABLE_HEADER_HEIGHT.name());// 得到表头占多少行
        // 一行一行填充
        for (int i = 0; i < jsonArray.size(); i++) {
            int rowNum = height + i;
            final Row[] row = {sheet.getRow(rowNum)};// 创建一行数据
            final JSONObject json = (JSONObject) jsonArray.get(i);// 得到这条数据
            AtomicReference<Exception> exception = new AtomicReference<>();
            tableBodyDesc.forEach((index, v) -> {
                if (row[0] == null) {
                    row[0] = sheet.createRow(rowNum);
                }
                // 给这个 index单元格 填入 value
                Cell cell = row[0].getCell(Integer.parseInt(index));// 得到单元格
                if (cell == null) {
                    cell = row[0].createCell(Integer.parseInt(index));
                    Workbook wb = sheet.getWorkbook();
                    Font font = sheet.getWorkbook().createFont();
                    font.setFontName("微软雅黑");
                    font.setFontHeightInPoints((short) 11);//设置字体大小
                    CellStyle cellStyle = wb.createCellStyle();
                    cellStyle.setLocked(true);
                    cellStyle.setFont(font);
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    cellStyle.setFillBackgroundColor((short) 12);
                    cell.setCellStyle(cellStyle);
                }

                if (v instanceof JSONArray) {
                    // 多个字段合并成一个单元格内容
                    JSONArray array = (JSONArray) v;
                    String[] linkedFormatString = new String[array.size()];
                    final AtomicInteger atomicInteger = new AtomicInteger(0);
                    array.forEach(obj -> {
                        // 处理每一个字段
                        JSONObject fieldDescData = (JSONObject) obj;
                        // 得到转换后的内容
                        final String resultValue = getFieldValue(json, fieldDescData);
                        linkedFormatString[atomicInteger.getAndIncrement()] = resultValue;// 存入该位置
                    });
                    final StringBuilder stringBuilder = new StringBuilder();
                    for (String s : linkedFormatString) {
                        stringBuilder.append(s);
                    }
                    String value = stringBuilder.toString();// 得到了合并后的内容
                    cell.setCellValue(value);

                } else {
                    // 一个字段可能要拆成多个单元格
                    JSONObject jsonObject = (JSONObject) v;
                    final String format = jsonObject.getString(ExcelCellEnum.FORMAT.name());
                    final Integer priority = jsonObject.getInteger(ExcelCellEnum.PRIORITY.name());
                    final String fieldName = jsonObject.getString(ExcelCellEnum.FIELD_NAME.name());
                    final String fieldType = jsonObject.getString(ExcelCellEnum.FIELD_TYPE.name());
                    final String size = jsonObject.getString(ExcelCellEnum.SIZE.name());
                    final String split = jsonObject.getString(ExcelCellEnum.SPLIT.name());

                    String fieldValue = String.valueOf(json.get(fieldName));// 得到这个字段值
                    final Integer width = jsonObject.getInteger(ExcelCellEnum.WIDTH.name());

                    if (width > 1) {
                        // 一个字段需要拆分成多个单元格
                        if (StringUtils.hasText(split)) {
                            // 有拆分词
                            final String[] splitArray = fieldValue.split(split);// 先拆分字段
                            if (StringUtils.hasText(format)) {
                                // 有格式化模板
                                final String[] formatStr = format.split(split);// 拆分后的格式化内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(Integer.parseInt(index) + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(Integer.parseInt(index) + j);// 得到单元格
                                    }
                                    String formattedStr = formatStr[j].replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            } else {
                                // 没有格式化模板直接填入内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(Integer.parseInt(index) + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(Integer.parseInt(index) + j);// 得到单元格
                                    }
                                    String formattedStr = format.replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            }

                        } else {
                            // 没有拆分词，本身需要拆分，抛异常
                            exception.set(new Exception(fieldName + "字段的注解上 缺少exportSplit拆分词"));
                        }
                    } else {
                        // 一个字段不需要拆成多个单元格
                        if (StringUtils.hasText(format)) {
                            // 内容存在格式化先进行格式化，然后填入值
                            String replacedStr = format.replace("$" + priority, fieldValue);// 替换字符串
                            cell.setCellValue(replacedStr);// 设置单元格内容
                        } else {
                            // 内容不需要格式化则直接填入(转换一下单位，如果没有就原样返回)
                            final String result = castForExport(fieldValue, fieldType, size);
                            cell.setCellValue(result);
                        }
                    }
                }
            });

            if (exception.get() != null) {
                throw exception.get();
            }
        }

    }

    /**
     * 得到对象的值(转换的只是格式化后的值)
     *
     * @param entity        实体数据转换的JSONObject
     * @param fieldDescData 字段规则描述数据
     * @return 处理后的字符串
     */
    private static String getFieldValue(JSONObject entity, JSONObject fieldDescData) {
        final String format = fieldDescData.getString(ExcelCellEnum.FORMAT.name());
        final Integer priority = fieldDescData.getInteger(ExcelCellEnum.PRIORITY.name());
        final String fieldName = fieldDescData.getString(ExcelCellEnum.FIELD_NAME.name());
        // 从对象中得到这个字段值
        String fieldValue = String.valueOf(entity.get(fieldName));// 得到这个字段值

        // 替换字符串
        return format.replace("$" + priority, fieldValue);
    }


    /**
     * 得到类型处理后的值字符串
     *
     * @param inputValue 输入的值
     * @param type       应该的类型
     * @param size       规模
     */
    public static String castForExport(String inputValue, String type, String size) {
        String result = "";
        if (!StringUtils.hasText(inputValue)) {
            return result;
        }
        final Class<?> fieldType = clazzMap.get(type);
        if (fieldType == String.class || fieldType == Character.class) {
            // 字符串或字符类型
            return inputValue;
        } else {
            // 是数值类型的进行转换
            BigDecimal bigDecimal = new BigDecimal(inputValue);
            final BigDecimal resultBigDecimal = BigDecimalUtils.bigDecimalDivBigDecimalFormatTwo(bigDecimal, new BigDecimal(size));
            if (fieldType == BigDecimal.class || fieldType == Double.class || fieldType == Float.class) {
                // 小数类型
                result = resultBigDecimal.toString();
            } else if (fieldType == Integer.class || fieldType == Long.class || fieldType == Short.class) {
                // 整数类型
                result = String.valueOf(resultBigDecimal.longValue());
            }
        }
        return result;
    }


    private static JSONArray sheetToJSONArray(JSONObject allExcelDescData, Sheet sheet) throws Exception {
        JSONArray result = new JSONArray();

        JSONObject tableHeader = getExcelHeaderDescInfo(allExcelDescData);// 表头描述信息
        JSONObject tableBody = getExcelBodyDescInfo(allExcelDescData);// 表的身体描述
        // 从表头描述信息得到表头的高
        Integer headerHeight = tableHeader.getInteger(TableEnum.TABLE_HEADER_HEIGHT.name());
        final int lastRowNum = sheet.getLastRowNum();
        boolean flag = false;// 是否是异常行
        String message = "";// 异常消息
        // 按行扫描excel表
        for (int i = headerHeight; i < lastRowNum; i++) {
            Row row = sheet.getRow(i);  // 得到第i行数据
            JSONObject oneRow = new JSONObject();// 一行数据
            int rowNum = i + 1;// 真正excel看到的行号
            oneRow.put(ExcelCellEnum.ROW.name(), rowNum);// 记录行号
            int length = tableBody.keySet().size();// 得到length,也就是需要转换的
            int count = 0;// 记录异常空字段次数，如果与length相等说明是空行
            //将Row转换为JSONObject
            for (Object entry : tableBody.values()) {
                JSONObject rowDesc = (JSONObject) entry;

                // 得到字段的索引位子
                Integer index = rowDesc.getInteger(ExcelCellEnum.COL.name());
                if (index < 0) continue;
                Integer width = rowDesc.getInteger(ExcelCellEnum.WIDTH.name());// 得到宽度，如果宽度不为1则需要进行合并多个单元格的内容

                String fieldName = rowDesc.getString(ExcelCellEnum.FIELD_NAME.name());// 字段名称
                String title = rowDesc.getString(ExcelCellEnum.TITLE_NAME.name());// 标题名称
                String fieldType = rowDesc.getString(ExcelCellEnum.FIELD_TYPE.name());// 字段类型
                String exception = rowDesc.getString(ExcelCellEnum.EXCEPTION.name());// 转换异常返回的消息
                String size = rowDesc.getString(ExcelCellEnum.SIZE.name());// 得到规模
                boolean nullable = rowDesc.getBoolean(ExcelCellEnum.NULLABLE.name());
                String positionMessage = "异常：第" + rowNum + "行的,第" + (index + 1) + "列 ---标题：" + title + " -- ";

                // 得到异常消息
                message = positionMessage + exception;

                // 获取合并的单元格值（合并后的结果，逗号分隔）
                String value = getMergeString(row, index, width);

                // 获取正则表达式，如果有正则，则进行正则截取value（相当于从单元格中取部分）
                String pattern = rowDesc.getString(ExcelCellEnum.PATTERN.name());
                boolean hasText = StringUtils.hasText(value);
                Object castValue = null;
                // 默认字段不可以为空，如果注解过程设置为true则不抛异常
                if (!nullable) {
                    // 说明字段不可以为空
                    if (!hasText) {
                        // 字段不能为空结果为空，这个空字段异常计数+1。除非count==length，然后重新计数，否则就是一行异常数据
                        count++;// 不为空字段异常计数+1
                        continue;
                    } else {
                        try {
                            // 单元格有内容,要么正常、要么异常直接抛不能返回null 直接中止
                            value = patternConvert(pattern, value);
                            castValue = cast(value, fieldType, message, size);
                        } catch (Exception e) {
                            throw new Exception(message);
                        }
                    }

                } else {
                    // 字段可以为空 （要么正常 要么返回null不会抛异常）
                    length--;
                    try {
                        // 单元格内容无关紧要。要么正常转换，要么返回null
                        value = patternConvert(pattern, value);
                        castValue = cast(value, fieldType, message, size);
                    } catch (Exception e) {
                        //castValue=null;// 本来初始值就是null
                    }
                }
                // 默认添加为null，只有正常才添加正常，否则中途抛异常直接中止
                oneRow.put(fieldName, castValue);// 添加数据
            }
            // 正常情况下count是等于length的，因为每个字段都需要处理
            if (count == 0) {
                result.add(oneRow);// 正常情况下添加一条数据
            } else if (count < length) {
                flag = true;// 需要抛异常，因为存在不合法数据
                break;// 非空行，并且遇到一行关键字段为null需要终止
            }
            // 空行继续扫描,或者正常

        }
        // 如果存在不合法数据抛异常
        if (flag) {
            throw new Exception(message);
        }

        return result;
    }


    /**
     * 返回Excel主体数据Body的描述信息
     */
    public static JSONObject getExcelBodyDescData(Class<T> clazz) {
        JSONObject partDescData = getDescriptionByClazz(clazz);
        return getExcelBodyDescInfo(partDescData);
    }

    /**
     * 从json中获取Excel身体部分数据
     */
    private static JSONObject getExcelBodyDescInfo(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(TableEnum.TABLE_BODY.name());
    }

    /**
     * 返回Excel头部Header的描述信息
     */
    public static JSONObject getExcelHeaderDescData(Class<T> clazz) {
        JSONObject partDescData = getDescriptionByClazz(clazz);
        return getExcelHeaderDescInfo(partDescData);
    }

    /**
     * 从json中获取Excel表头部分数据
     */
    private static JSONObject getExcelHeaderDescInfo(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(TableEnum.TABLE_HEADER.name());
    }

    /**
     * 传入Sheet获取一个完整的表格描述信息，将INDEX更新
     *
     * @param excelDescData excel的描述信息
     * @param sheet         待解析的excel表格
     */
    private static JSONObject filledTitleIndexBySheet(JSONObject excelDescData, Sheet sheet) {
        JSONObject tableHeaderDesc = getExcelHeaderDescInfo(excelDescData);
        JSONObject tableBodyDesc = getExcelBodyDescInfo(excelDescData);
        Integer height = tableHeaderDesc.getInteger(TableEnum.TABLE_HEADER_HEIGHT.name());// 得到表头占据了那几行

        // 补充table每一项的index信息
        tableBodyDesc.forEach((fieldName, cellDesc) -> {
            // 扫描包含表头的那几行 记录需要记录的标题所在的索引列，填充INDEX
            for (int i = 0; i < height; i++) {
                Row row = sheet.getRow(i);// 得到第i行数据（在表头内）
                // 遍历这行所有单元格，然后得到表头进行比较找到标题和注解上的titleName相同的单元格
                row.forEach(cell -> {
                    // 得到单元格内容（统一为字符串类型）
                    String title = getStringCellValue(cell);

                    JSONObject cd = (JSONObject) cellDesc;
                    // 如果标题相同找到了这单元格，获取单元格下标存入
                    if (title.equals(cd.getString(ExcelCellEnum.TITLE_NAME.name()))) {
                        int columnIndex = cell.getColumnIndex();// 找到了则取出索引存入jsonObject
                        cd.put(ExcelCellEnum.COL.name(), columnIndex); // 补全描述信息
                    }
                });
            }

        });
        System.out.println(excelDescData);
        return excelDescData;
    }

    /**
     * 获取完整的Excel描述信息
     *
     * @param tClass 模板类
     * @param sheet  Excel
     * @param <T>    泛型
     */
    private static <T> JSONObject getFulledExcelDescriptionInfo(Class<T> tClass, Sheet sheet) {
        // 获取表格部分描述信息（根据泛型得到的）
        JSONObject partDescData = getDescriptionByClazz(tClass);
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
    private static JSONObject getDescriptionByClazz(Class<?> clazz) {
        Field[] fields = clazz.getDeclaredFields();// 获取所有字段
        JSONObject excelDescData = new JSONObject();// excel的描述数据
        JSONObject tableBody = new JSONObject();// 表中主体数据信息
        JSONObject tableHeader = new JSONObject();// 表中主体数据信息

        // 1、先得到表头信息
        final ExcelTableHeader tableHeaderAnnotation = clazz.getAnnotation(ExcelTableHeader.class);
        if (tableHeaderAnnotation != null) {
            tableHeader.put(TableEnum.TABLE_NAME.name(), tableHeaderAnnotation.tableName());// 表的名称
            tableHeader.put(TableEnum.TABLE_HEADER_HEIGHT.name(), tableHeaderAnnotation.height());// 表头的高度

            // 2、得到表的Body信息
            for (Field field : fields) {
                final ExcelCellBind annotationTitle = field.getDeclaredAnnotation(ExcelCellBind.class);
                if (annotationTitle != null) {// 找到自定义的注解
                    JSONObject cellDesc = new JSONObject();// 单元格描述信息
                    String title = annotationTitle.title();         // 获取标题，如果标题不存在则不进行处理
                    if (StringUtils.hasText(title)) {

                        cellDesc.put(ExcelCellEnum.TITLE_NAME.name(), title);// 标题名称
                        cellDesc.put(ExcelCellEnum.FIELD_NAME.name(), field.getName());// 字段名称
                        cellDesc.put(ExcelCellEnum.FIELD_TYPE.name(), field.getType().getTypeName());// 字段的类型
                        cellDesc.put(ExcelCellEnum.WIDTH.name(), annotationTitle.width());// 单元格的宽度（宽度为2代表合并了2格单元格）
                        cellDesc.put(ExcelCellEnum.EXCEPTION.name(), annotationTitle.exception());// 校验如果失败返回的异常消息
                        cellDesc.put(ExcelCellEnum.COL.name(), annotationTitle.index());// 默认的索引位置
                        cellDesc.put(ExcelCellEnum.SIZE.name(), annotationTitle.size());// 规模,记录规模(亿元/万元)
                        cellDesc.put(ExcelCellEnum.PATTERN.name(), annotationTitle.importPattern());// 正则表达式
                        cellDesc.put(ExcelCellEnum.NULLABLE.name(), annotationTitle.nullable());// 是否可空
                        cellDesc.put(ExcelCellEnum.SPLIT.name(), annotationTitle.exportSplit());// 导出字段的拆分
                        cellDesc.put(ExcelCellEnum.FORMAT.name(), annotationTitle.exportFormat());// 导出的模板格式
                        cellDesc.put(ExcelCellEnum.PRIORITY.name(), annotationTitle.exportPriority());// 导出拼串的顺序

                        // 以字段名作为key
                        tableBody.put(field.getName(), cellDesc);// 存入这个标题名单元格的的描述信息，后面还需要补全INDEX
                    }
                }
            }
        }
        excelDescData.put(TableEnum.TABLE_HEADER.name(), tableHeader);// 将表头记录信息注入
        excelDescData.put(TableEnum.TABLE_BODY.name(), tableBody);// 将表的body记录信息注入
        return excelDescData;// 返回记录的所有信息
    }


    /**
     * 获取导出所需要的所有信息
     *
     * @param clazz 传入的泛型，注解信息
     * @param sheet excel表格
     * @return 表格的信息
     */
    public static JSONObject getExportDescriptionByClazzAndSheet(Class<?> clazz, Sheet sheet) {
        Field[] fields = clazz.getDeclaredFields();// 获取所有字段
        JSONObject tableHeader = new JSONObject();// 表中主体数据信息
        JSONObject tableBody = new JSONObject();// 表中主体数据信息
        JSONObject excelDescriptionInfo = new JSONObject();// excel的描述数据

        // 1、先得到表头信息
        final ExcelTableHeader tableHeaderAnnotation = clazz.getDeclaredAnnotation(ExcelTableHeader.class);
        if (tableHeaderAnnotation != null) {
            tableHeader.put(TableEnum.TABLE_NAME.name(), tableHeaderAnnotation.tableName());// 表的名称
            tableHeader.put(TableEnum.TABLE_HEADER_HEIGHT.name(), tableHeaderAnnotation.height());// 表头的高度
            tableHeader.put(TableEnum.RESOURCE.name(), tableHeaderAnnotation.resource());// 模板excel的访问路径
            tableHeader.put(TableEnum.PASSWORD.name(), tableHeaderAnnotation.password());// 模板excel的访问路径

            // 2、得到表的Body信息
            for (Field field : fields) {
                final ExcelCellBind annotationTitle = field.getDeclaredAnnotation(ExcelCellBind.class);// 获取ExcelCellEnumBindAnnotation注解
                ExcelCellStyle annotationStyle = field.getDeclaredAnnotation(ExcelCellStyle.class);// 获取单元格样式注解
                if (annotationTitle != null) {// 找到自定义的注解
                    JSONObject titleDesc = new JSONObject();// 单元格描述信息
                    String title = annotationTitle.title();         // 获取标题，如果标题不存在则不进行处理
                    if (StringUtils.hasText(title)) {
                        // 根据标题找到索引
                        int titleIndexFromSheet = getTitleIndexFromSheet(title, sheet, tableHeaderAnnotation.height());
                        String titleIndexString = "" + titleIndexFromSheet;// 字符串类型 标题的下标

                        titleDesc.put(ExcelCellEnum.WIDTH.name(), annotationTitle.width());// 单元格的宽度（宽度为2代表合并了2格单元格）
                        titleDesc.put(ExcelCellEnum.SPLIT.name(), annotationTitle.exportSplit());// 导出字符串拆分为多个单元格（州市+区县）
                        titleDesc.put(ExcelCellEnum.FIELD_NAME.name(), field.getName());// 导出字符串拆分为多个单元格（州市+区县）
                        titleDesc.put(ExcelCellEnum.FORMAT.name(), annotationTitle.exportFormat());// 这个字段输出的格式
                        titleDesc.put(ExcelCellEnum.PRIORITY.name(), annotationTitle.exportPriority());// 这个字段输出的格式
                        titleDesc.put(ExcelCellEnum.SIZE.name(), annotationTitle.size());// 字段规模
                        titleDesc.put(ExcelCellEnum.FIELD_TYPE.name(), field.getType().getTypeName());// 字段类型
                        String fontName = "微软雅黑";
                        short fontSize = 11, bgColor = 9, rotation = 0, foregroundColor = 9;
                        boolean locked = false, hidden = false, wrapText = false, shrink = false;
                        HorizontalAlignment textAlign = HorizontalAlignment.CENTER;
                        VerticalAlignment verticalAlignment = VerticalAlignment.CENTER;
                        FillPatternType fillPatternType = FillPatternType.NO_FILL;
                        BorderStyle top = BorderStyle.NONE, bottom = BorderStyle.NONE, left = BorderStyle.NONE, right = BorderStyle.NONE;
                        if (annotationStyle != null) {
                            fontName = annotationStyle.fontName();
                            fontSize = annotationStyle.fontSize();
                            textAlign = annotationStyle.textAlign();
                            bgColor = annotationStyle.backgroundColor();
                            locked = annotationStyle.locked();
                            hidden = annotationStyle.hidden();
                            wrapText = annotationStyle.wrapText();
                            verticalAlignment = annotationStyle.verticalAlign();
                            rotation = annotationStyle.rotation();
                            fillPatternType = annotationStyle.fillPatternType();
                            foregroundColor = annotationStyle.foregroundColor();
                            shrink = annotationStyle.autoShrink();
                            top = annotationStyle.top();
                            bottom = annotationStyle.bottom();
                            left = annotationStyle.left();
                            right = annotationStyle.right();
                        }
                        titleDesc.put(CellStyleEnum.FONT_NAME.name(), fontName);// 字体名称
                        titleDesc.put(CellStyleEnum.FONT_SIZE.name(), fontSize);// 字体大小
                        titleDesc.put(CellStyleEnum.TEXT_ALIGN.name(), textAlign);// 文本位置，居中什么的
                        titleDesc.put(CellStyleEnum.BG_COLOR.name(), bgColor);// 背景颜色
                        titleDesc.put(CellStyleEnum.LOCKED.name(), locked);// 是否可编辑
                        titleDesc.put(CellStyleEnum.HIDDEN.name(), hidden);// 是否启用样式
                        titleDesc.put(CellStyleEnum.WRAP_TEXT.name(), wrapText);// 多行显示
                        titleDesc.put(CellStyleEnum.VERTICAL_ALIGN.name(), verticalAlignment);// 垂直对齐
                        titleDesc.put(CellStyleEnum.ROTATION.name(), rotation);// 旋转
                        titleDesc.put(CellStyleEnum.FILL_PATTERN.name(), fillPatternType);// 填充图案
                        titleDesc.put(CellStyleEnum.FORE_COLOR.name(), foregroundColor);// 前景颜色
                        titleDesc.put(CellStyleEnum.AUTO_SHRINK.name(), shrink);// 自动缩小
                        titleDesc.put(CellStyleEnum.TOP.name(), top);// 上边框样式
                        titleDesc.put(CellStyleEnum.BOTTOM.name(), bottom);//下
                        titleDesc.put(CellStyleEnum.LEFT.name(), left);//左
                        titleDesc.put(CellStyleEnum.RIGHT.name(), right);//右
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
                                tableBody.put(titleIndexString, array);// JSONObject转化为JSONArray
                            }
                        } else {
                            // 没有相同的title则直接添加
                            tableBody.put(titleIndexString, titleDesc);
                        }
                    }
                }
            }
        }
        excelDescriptionInfo.put(TableEnum.TABLE_HEADER.name(), tableHeader);// 将表头记录信息注入
        excelDescriptionInfo.put(TableEnum.TABLE_BODY.name(), tableBody);// 将表的body记录信息注入
        System.out.println(excelDescriptionInfo);
        return excelDescriptionInfo;// 返回记录的所有信息
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
            Row row = sheet.getRow(i);// 得到第i行数据（在表头内）
            // 遍历这行所有单元格，然后得到表头进行比较找到标题和注解上的titleName相同的单元格
            for (Cell cell : row) {// 得到单元格内容（统一为字符串类型）
                String titleName = getStringCellValue(cell);
                // 如果标题相同找到了这单元格，获取单元格下标存入
                if (title.equals(titleName)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return index;
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
    private static Object cast(String inputValue, Class<?> aClass, String exception, String size) throws
            ClassCastException {
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
            throw new ClassCastException("类型转换异常，输入的文本内容：" + inputValue);
            //e.printStackTrace();
        }

        return obj;
    }

}
