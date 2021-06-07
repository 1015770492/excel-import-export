package top.yumbo.excel.util;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.util.StringUtils;
import org.springframework.web.client.RestTemplate;
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelCellStyle;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.entity.CellStyleBuilder;

import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author jinhua
 * @date 2021/5/21 21:51
 * <p>
 * 整体的设计思路：
 * 1、导入：list->jsonArray->sheet
 * 2、导出：sheet->jsonArray->list
 * 转换中需要用到的信息通过注解实现，
 * 对于导入：通过注解-> 得到字段和表格关系
 * 对于导出：通过注解-> 得到表头和表头关系
 * 自定义样式：使用java8的函数式编程思想，将设定样式的功能通过函数式接口抽离
 * 通过外部编码使样式的设置更加方便
 */
public class ExcelImportExportUtils {



    //表头信息
    public enum TableEnum {
        WORK_BOOK, SHEET, TABLE_NAME, TABLE_HEADER, TABLE_HEADER_HEIGHT, RESOURCE, TYPE, TABLE_BODY, PASSWORD
    }

    // 单元格信息
    public enum CellEnum {
        TITLE_NAME, FIELD_NAME, FIELD_TYPE, SIZE, PATTERN, NULLABLE, WIDTH, EXCEPTION, COL, ROW, SPLIT, PRIORITY, FORMAT
    }

    //样式的属性名
    public enum CellStyleEnum {
        FONT_NAME, FONT_SIZE, BG_COLOR, TEXT_ALIGN, LOCKED, HIDDEN, BOLD,
        VERTICAL_ALIGN, WRAP_TEXT, STYLES,
        FORE_COLOR, ROTATION, FILL_PATTERN, AUTO_SHRINK, TOP, BOTTOM, LEFT, RIGHT
    }


    /**
     * 导入Excel数据
     *
     * @param workbook excel工作簿
     * @param tClass   泛型
     */
    public static <T> List<T> importExcel(Workbook workbook, Class<T> tClass) throws Exception {
        return excelToList(workbook.getSheetAt(0), tClass);
    }

    /**
     * 导入excel，默认导入xlsx文件
     *
     * @param inputStream 输入流
     * @param tClass      泛型
     */
    public static <T> List<T> importExcel(InputStream inputStream, Class<T> tClass) throws Exception {
        return importExcelForXlsx(inputStream, tClass);
    }

    /**
     * 导入xlsx的文件
     *
     * @param inputStream 输入流
     * @param tClass      泛型
     */
    public static <T> List<T> importExcelForXlsx(InputStream inputStream, Class<T> tClass) throws Exception {
        return importExcel(inputStream, tClass, "xlsx");
    }

    /**
     * 导入xls文件
     *
     * @param inputStream 输入流
     * @param tClass      泛型
     */
    public static <T> List<T> importExcelForXls(InputStream inputStream, Class<T> tClass) throws Exception {
        return importExcel(inputStream, tClass, "xls");
    }

    /**
     * 导入excel
     *
     * @param inputStream 输入流
     * @param type        类型：xls、xlsx
     * @param tClass      模板数据
     * @return List类型的实体
     */
    public static <T> List<T> importExcel(InputStream inputStream, Class<T> tClass, String type) throws Exception {
        if (type == null) {
            type = "xlsx";
        }
        Workbook workbook = null;
        if ("xls".equals(type)) {
            workbook = new HSSFWorkbook(inputStream);
        } else if ("xlsx".equals(type)) {
            workbook = new XSSFWorkbook(inputStream);
        }
        if (workbook != null) {
            return importExcel(workbook, tClass);
        }
        return null;
    }

    /**
     * 导出Excel,使用默认样式  （传入list 和 输出流）
     *
     * @param list         待导入的数据集合
     * @param outputStream 导出的文件输出流
     */
    public static <T> void exportExcel(List<T> list, OutputStream outputStream) throws Exception {
        if (list != null && list.size() > 0 && outputStream != null) {
            // 根据泛型得到导出的信息以及模板
            final JSONObject exportInfo = getExportInfo(list.get(0).getClass());
            final Sheet sheet = getSheet(exportInfo);
            // 得到工作簿对象
            final Workbook workbook = getWorkBook(exportInfo);
            // 将数据填入到表格（默认样式）
            listToExcel(list, sheet);
            workbook.write(outputStream);
            System.out.println("默认样式导出完毕");
            workbook.close();
        } else if (list == null) {
            throw new NullPointerException("list不能为null");
        } else {
            throw new NullPointerException("输出流不能为空");
        }
    }


    /**
     * 行高亮显示
     *
     * @param list         数据
     * @param outputStream 导出的文件的输出流
     * @param function     功能性函数，返回颜色值
     */
    public static <T> void exportExcelRowHighLight(List<T> list, OutputStream outputStream, Function<T, IndexedColors> function) throws Exception {
        final JSONObject exportInfo = getExportInfo(list.get(0).getClass());
        final JSONObject titleInfo = getExcelBodyDescInfo(exportInfo);
        final JSONArray jsonArray = listToJSONArray(list);
        final Sheet sheet = getSheet(exportInfo);
        final Integer height = getTableHeight(getExcelHeaderDescInfo(exportInfo));

        // 一行一行填充
        for (int i = 0; i < jsonArray.size(); i++) {
            int rowNum = height + i;
            final Row[] row = {sheet.getRow(rowNum)};// 创建一行数据
            final JSONObject json = jsonArray.getJSONObject(i);
            AtomicReference<Exception> exception = new AtomicReference<>();
            // 遍历表身体信息
            int finalI = i;
            titleInfo.forEach((titleIdx, v) -> {
                if (row[0] == null) {
                    row[0] = sheet.createRow(rowNum);
                }
                // 标题 索引
                int titleIndex = Integer.parseInt(titleIdx);
                if (titleIndex < 0) {
                    return;
                }
                // 给这个 index单元格 填入 value
                Cell cell = row[0].getCell(titleIndex);// 得到单元格
                if (cell == null) {
                    cell = row[0].createCell(titleIndex);
                }
                Workbook wb = sheet.getWorkbook();
                CellStyle cellStyle = wb.createCellStyle();
                Font font = sheet.getWorkbook().createFont();
                font.setFontName("微软雅黑");
                font.setFontHeightInPoints((short) 11);//设置字体大小
                cellStyle.setLocked(true);
                cellStyle.setFont(font);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                final short index = function.apply(list.get(finalI)).getIndex();
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setFillForegroundColor(index);// 设置颜色
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
                    String fieldValue = String.valueOf(json.get(fieldName));// 得到这个字段值
                    final Integer width = jsonObject.getInteger(CellEnum.WIDTH.name());

                    if (width > 1) {
                        // 一个字段需要拆分成多个单元格
                        if (StringUtils.hasText(split)) {
                            // 有拆分词
                            final String[] splitArray = fieldValue.split(split);// 先拆分字段
                            if (StringUtils.hasText(format)) {
                                // 有格式化模板
                                final String[] formatStr = format.split(split);// 拆分后的格式化内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(titleIndex + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(titleIndex + j);// 得到单元格
                                    }
                                    String formattedStr = formatStr[j].replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellStyle(cellStyle);
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            } else {
                                // 没有格式化模板直接填入内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(titleIndex + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(titleIndex + j);// 得到单元格
                                    }
                                    String formattedStr = format.replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellStyle(cellStyle);
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
        getWorkBook(exportInfo).write(outputStream);
        System.out.println("高亮行显示导出完毕");
    }
    public static <T> void exportExcelRowHighLightRGBColor(List<T> list, FileOutputStream fileOutputStream, Function<T,Color> function) {
        byte []rgb={new Integer(212).byteValue(),44,new Integer(144).byteValue()};
        final XSSFColor xssfColor = new XSSFColor(rgb, null);

    }

    /**
     * (导入)将sheet解析成List类型的数据
     * （注意这里只是将单元格内容转换为了实体，具体字段可能还不是正确的例如 区域码应该是是具体的编码而不是XX市XX区）
     *
     * @param tClass 传入的泛型
     * @param sheet  表单数据（带表头的）
     * @return 只是将单元格内容转化为List
     */
    private static <T> List<T> excelToList(Sheet sheet, Class<T> tClass) throws Exception {
        JSONArray jsonArray = excelToJSONArray(sheet, tClass);
        return JSONArray.parseArray(jsonArray.toJSONString(), tClass);
    }

    /**
     * （导入）将excel表转换为JSONArray
     *
     * @param tClass 注解模板类
     * @param sheet  传入的excel数据
     */
    private static <T> JSONArray excelToJSONArray(Sheet sheet, Class<T> tClass) throws Exception {
        JSONObject fulledExcelDescData = getImportInfoByClazz(sheet, tClass);
        // 根据所有已知信息将excel转换为JsonArray数据
        return sheetToJSONArray(fulledExcelDescData, sheet);
    }

    /**
     * （导出）List转换为表格
     *
     * @param list  数据集
     * @param sheet 待填入的excel表格
     * @return excel工作簿
     * @throws Exception 抛出的异常
     */
    private static <T> void listToExcel(List<T> list, Sheet sheet) throws Exception {
        if (list != null && sheet != null) {
            if (list.size() > 0) {
                final JSONArray listArray = listToJSONArray(list);
                // 获取导出的描述信息
                final JSONObject exportInfo = getExportInfo(list.get(0).getClass(), sheet);

                // 将list按照导出的描述信息填入Excel
                jsonArrayToExcel(listArray, exportInfo, sheet);

            } else {
                throw new Exception("list集合没有数据");
            }

        } else if (list == null) {
            throw new NullPointerException("list不能为空");
        } else {
            throw new NullPointerException("sheet不能为空");
        }
    }


    /**
     * （导出）jsonArray转换为Sheet
     * 存在模板的情况下
     * 将数据填充进入Excel表格
     *
     * @param jsonArray  数据
     * @param exportInfo 规则
     * @param sheet      excel表格
     */
    private static void jsonArrayToExcel(JSONArray jsonArray, JSONObject exportInfo, Sheet sheet) throws Exception {
        // 得到表头信息
        JSONObject headerInfo = getExcelHeaderDescInfo(exportInfo);
        // 得到表的身体描述信息
        JSONObject titleInfo = getExcelBodyDescInfo(exportInfo);
        // 从表头信息得到表头的高度
        final Integer height = getTableHeight(headerInfo);
        // 一行一行填充
        for (int i = 0; i < jsonArray.size(); i++) {
            int rowNum = height + i;
            final Row[] row = {sheet.getRow(rowNum)};// 创建一行数据
            final JSONObject json = jsonArray.getJSONObject(i);
            AtomicReference<Exception> exception = new AtomicReference<>();
            // 遍历表身体信息
            titleInfo.forEach((titleIdx, v) -> {
                if (row[0] == null) {
                    row[0] = sheet.createRow(rowNum);
                }
                // 标题 索引
                int titleIndex = Integer.parseInt(titleIdx);
                if (titleIndex < 0) {
                    return;
                }
                // 给这个 index单元格 填入 value
                Cell cell = row[0].getCell(titleIndex);// 得到单元格
                if (cell == null) {
                    cell = row[0].createCell(titleIndex);
                    Workbook wb = sheet.getWorkbook();
                    Font font = sheet.getWorkbook().createFont();
                    font.setFontName("微软雅黑");
                    font.setFontHeightInPoints((short) 11);//设置字体大小
                    CellStyle cellStyle = wb.createCellStyle();
                    cellStyle.setLocked(true);
                    cellStyle.setFont(font);
                    cellStyle.setAlignment(HorizontalAlignment.CENTER);
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cellStyle.setFillForegroundColor((short) 12);
                    cell.setCellStyle(cellStyle);
                }

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
                    String fieldValue = String.valueOf(json.get(fieldName));// 得到这个字段值
                    final Integer width = jsonObject.getInteger(CellEnum.WIDTH.name());

                    if (width > 1) {
                        // 一个字段需要拆分成多个单元格
                        if (StringUtils.hasText(split)) {
                            // 有拆分词
                            final String[] splitArray = fieldValue.split(split);// 先拆分字段
                            if (StringUtils.hasText(format)) {
                                // 有格式化模板
                                final String[] formatStr = format.split(split);// 拆分后的格式化内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(titleIndex + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(titleIndex + j);// 得到单元格
                                    }
                                    String formattedStr = formatStr[j].replace("$" + j, splitArray[j]);// 替换字符串
                                    cell.setCellValue(formattedStr);// 将格式化后的字符串填入
                                }
                            } else {
                                // 没有格式化模板直接填入内容
                                for (int j = 0; j < width; j++) {
                                    cell = row[0].getCell(titleIndex + j);
                                    if (cell == null) {
                                        cell = row[0].createCell(titleIndex + j);// 得到单元格
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
     * 高亮显示某一个标题列并且进行断言
     */
    private static <T> Workbook listToExcel(List<T> list, Sheet sheet, Function<T, String> function) throws Exception {
        if (list != null && list.size() > 0 && sheet != null && function != null) {
            // 获取导出的描述信息
            final JSONObject exportInfo = getExportInfo(list.get(0).getClass(), sheet);
            // 将list按照导出的描述信息填入Excel
            jsonArrayToExcel(listToJSONArray(list), exportInfo, sheet);

            return getWorkBook(exportInfo);
        } else if (list == null) {
            throw new NullPointerException("list不能为空");
        } else {
            throw new NullPointerException("sheet不能为空");
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
     * list转JSONArray
     * 返回 字段名称 -> 值 的键值对，后面根据字段名称得到值
     *
     * @param list 集合
     */
    private static JSONArray listToJSONArray(List<?> list) {
        return JSONObject.parseArray(JSONObject.toJSONString(list));
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
     * 得到类型处理后的值字符串
     *
     * @param inputValue 输入的值
     * @param type       应该的类型
     * @param size       规模
     */
    private static String castForExport(String inputValue, String type, String size) {
        String result = "";
        if (!StringUtils.hasText(inputValue)) {
            return result;
        }
        final Class<?> fieldType = clazzMap.get(type);
        if (fieldType == String.class || fieldType == Character.class) {
            // 字符串或字符类型
            return inputValue;
        } else if (fieldType == Date.class || fieldType == LocalDate.class
                || fieldType == LocalDateTime.class || fieldType == LocalTime.class) {
            // 时间类型的

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


    /**
     * 转换为excel数据转换为 JSONArray
     */
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
            oneRow.put(CellEnum.ROW.name(), rowNum);// 记录行号
            int length = tableBody.keySet().size();// 得到length,也就是需要转换的
            int count = 0;// 记录异常空字段次数，如果与length相等说明是空行
            //将Row转换为JSONObject
            for (Object entry : tableBody.values()) {
                JSONObject rowDesc = (JSONObject) entry;

                // 得到字段的索引位子
                Integer index = rowDesc.getInteger(CellEnum.COL.name());
                if (index < 0) continue;
                Integer width = rowDesc.getInteger(CellEnum.WIDTH.name());// 得到宽度，如果宽度不为1则需要进行合并多个单元格的内容

                String fieldName = rowDesc.getString(CellEnum.FIELD_NAME.name());// 字段名称
                String title = rowDesc.getString(CellEnum.TITLE_NAME.name());// 标题名称
                String fieldType = rowDesc.getString(CellEnum.FIELD_TYPE.name());// 字段类型
                String exception = rowDesc.getString(CellEnum.EXCEPTION.name());// 转换异常返回的消息
                String size = rowDesc.getString(CellEnum.SIZE.name());// 得到规模
                boolean nullable = rowDesc.getBoolean(CellEnum.NULLABLE.name());
                String positionMessage = "异常：第" + rowNum + "行的,第" + (index + 1) + "列 ---标题：" + title + " -- ";

                // 得到异常消息
                message = positionMessage + exception;

                // 获取合并的单元格值（合并后的结果，逗号分隔）
                String value = getMergeString(row, index, width);

                // 获取正则表达式，如果有正则，则进行正则截取value（相当于从单元格中取部分）
                String pattern = rowDesc.getString(CellEnum.PATTERN.name());
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
                            throw new Exception(message + e.getMessage());
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
    private static JSONObject getExcelBodyDescData(Class<T> clazz) {
        JSONObject partDescData = getImportDescriptionByClazz(clazz);
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
    private static JSONObject getExcelHeaderDescData(Class<T> clazz) {
        JSONObject partDescData = getImportDescriptionByClazz(clazz);
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
        Integer height = getTableHeight(tableHeaderDesc);// 得到表头占据了那几行

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
                    if (title.equals(cd.getString(CellEnum.TITLE_NAME.name()))) {
                        int columnIndex = cell.getColumnIndex();// 找到了则取出索引存入jsonObject
                        cd.put(CellEnum.COL.name(), columnIndex); // 补全描述信息
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
    private static <T> JSONObject getImportInfoByClazz(Sheet sheet, Class<T> tClass) {
        // 获取表格部分描述信息（根据泛型得到的）
        JSONObject partDescData = getImportDescriptionByClazz(tClass);
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
        excelDescData.put(TableEnum.TABLE_BODY.name(), tableBody);// 将表的body记录信息注入
        return excelDescData;// 返回记录的所有信息
    }


    /**
     * 获取导出的信息
     *
     * @param clazz 泛型
     */
    private static JSONObject getExportInfo(Class<?> clazz) throws Exception {
        return getExportInfo(clazz, null);
    }

    /**
     * 获取导出所需要的所有信息（标题-字段关系）
     * 注意：会将样式也创建好并且放入返回的jsonObject中，然后直接取
     *
     * @param clazz 传入的泛型，注解信息
     * @param sheet excel表格
     * @return 表格的信息，返回的tableBody中key是每一个标题的索引，value则是由字段信息
     */
    private static JSONObject getExportInfo(Class<?> clazz, Sheet sheet) throws IOException, IllegalArgumentException {
        Field[] fields = clazz.getDeclaredFields();// 获取所有字段
        JSONObject tableHeader = new JSONObject();// 表中主体数据信息
        JSONObject tableBody = new JSONObject();// 表中主体数据信息
        JSONObject exportInfo = new JSONObject();// excel的描述数据

        // 1、先得到表头信息
        final ExcelTableHeader tableHeaderAnnotation = clazz.getDeclaredAnnotation(ExcelTableHeader.class);
        if (tableHeaderAnnotation != null) {
            tableHeader.put(TableEnum.TABLE_NAME.name(), tableHeaderAnnotation.tableName());// 表的名称
            tableHeader.put(TableEnum.TABLE_HEADER_HEIGHT.name(), tableHeaderAnnotation.height());// 表头的高度
            tableHeader.put(TableEnum.RESOURCE.name(), tableHeaderAnnotation.resource());// 模板excel的访问路径
            tableHeader.put(TableEnum.TYPE.name(), tableHeaderAnnotation.type());// xlsx 或 xls 格式
            tableHeader.put(TableEnum.PASSWORD.name(), tableHeaderAnnotation.password());// 模板excel的访问路径
            if (sheet == null) {
                // sheet不存在则从注解信息中获取
                final Workbook workBook = getWorkBookByResource(tableHeaderAnnotation.resource(), tableHeaderAnnotation.type());
                sheet = workBook.getSheetAt(0);
                if (sheet == null) {
                    sheet = workBook.createSheet();
                }
            }
            exportInfo.put(TableEnum.SHEET.name(), sheet);
            exportInfo.put(TableEnum.WORK_BOOK.name(), sheet.getWorkbook());

            // 2、得到表的Body信息
            for (Field field : fields) {
                final ExcelCellBind annotationTitle = field.getDeclaredAnnotation(ExcelCellBind.class);// 获取ExcelCellEnumBindAnnotation注解
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
    private static Workbook getWorkBookByResource(String resourcePath, String type) throws IOException, IllegalArgumentException {
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
                    try (final InputStream is = new ByteArrayInputStream(body)) {//java9新特性try升级 自动关闭流
                        inputStream = is;
                    }
                } else {
                    final String[] split = resourcePath.split("://");
                    if (split[1].startsWith("/")) {
                        // 是相对路径，给他转为绝对路径
                        File directory = new File("");//设定为当前文件夹
                        String currentAbsolutePath = directory.getAbsolutePath();
                        resourcePath = currentAbsolutePath + split[1];// 得到新的绝对路径
                    } else {
                        resourcePath = split[1];
                    }
                    // 绝对路径
                    inputStream = new FileInputStream(resourcePath);
                }
                if ("xlsx".equals(type)) {
                    return new XSSFWorkbook(inputStream);
                } else {
                    return new HSSFWorkbook(inputStream);
                }
            }
            throw new IllegalArgumentException("请带上协议头例如http://");
        } else {
            throw new IllegalArgumentException("资源地址不正确，配置的资源地址：" + resourcePath);
        }
    }

    /**
     * 设置单元格样式，注解上注入的，如果注解上为空还会有一个默认的样式
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
        final boolean bold = cellStyle.getBoolean(CellStyleEnum.BOLD.name());
        final boolean shrink = cellStyle.getBoolean(CellStyleEnum.AUTO_SHRINK.name());

        final HorizontalAlignment textAlign = cellStyle.getObject(CellStyleEnum.TEXT_ALIGN.name(), HorizontalAlignment.class);
        final VerticalAlignment verticalAlignment = cellStyle.getObject(CellStyleEnum.VERTICAL_ALIGN.name(), VerticalAlignment.class);
        final FillPatternType fillPatternType = cellStyle.getObject(CellStyleEnum.FILL_PATTERN.name(), FillPatternType.class);
        final BorderStyle top = cellStyle.getObject(CellStyleEnum.TOP.name(), BorderStyle.class);
        final BorderStyle bottom = cellStyle.getObject(CellStyleEnum.BOTTOM.name(), BorderStyle.class);
        final BorderStyle left = cellStyle.getObject(CellStyleEnum.LEFT.name(), BorderStyle.class);
        final BorderStyle right = cellStyle.getObject(CellStyleEnum.RIGHT.name(), BorderStyle.class);
        setCellStyle(cell, fontName, fontSize, bold, locked, hidden, textAlign, bgColor, foreColor, rotation, verticalAlignment, fillPatternType, top, bottom, left, right, wrapText, shrink);
    }

    private static void setCellStyle(Cell cell,
                                     String fontName, Short fontSize, Boolean bold, Boolean locked, Boolean hidden,
                                     HorizontalAlignment textAlign,
                                     Short bgColor, Short foreignColor,
                                     Short rotation, VerticalAlignment verticalAlignment, FillPatternType fillPatternType,
                                     BorderStyle top, BorderStyle bottom, BorderStyle left, BorderStyle right, Boolean wrapText,
                                     Boolean autoShrink
    ) {

        CellStyle cellStyle = getCellStyle(cell.getSheet().getWorkbook(), fontName, (int) fontSize, bold, locked, hidden, textAlign, (int) bgColor, (int) foreignColor, (int) rotation, verticalAlignment, fillPatternType, top, bottom, left, right, wrapText, autoShrink);
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
    private static CellStyle getCellStyle(Workbook workbook, String fontName, Integer fontSize, Boolean bold, Boolean locked, Boolean hidden, HorizontalAlignment textAlign,
                                          Integer bgColor, Integer foreignColor,
                                          Integer rotation, VerticalAlignment verticalAlignment, FillPatternType fillPatternType,
                                          BorderStyle top, BorderStyle bottom, BorderStyle left, BorderStyle right, Boolean wrapText,
                                          Boolean autoShrink
    ) {
        return CellStyleBuilder.builder().fontName(fontName).fontSize(fontSize).bold(bold).locked(locked)
                .hidden(hidden).textAlign(textAlign).bgColor(bgColor).foregroundColor(foreignColor)
                .rotation(rotation).verticalAlignment(verticalAlignment)
                .fillPatternType(fillPatternType)
                .top(top).bottom(bottom).left(left).right(right).wrapText(wrapText)
                .autoShrink(autoShrink).build().getCellStyle(workbook);
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
            throw new ClassCastException("类型转换异常，输入的文本内容（=><=符合中间就是待转换的内容）：=>" + inputValue + "<=.位置：" + exception);
        }

        return obj;
    }

}
