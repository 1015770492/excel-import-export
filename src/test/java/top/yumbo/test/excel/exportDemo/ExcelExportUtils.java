package top.yumbo.test.excel.exportDemo;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import lombok.Builder;
import lombok.Data;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
import top.yumbo.excel.annotation.ExcelCellBind;
import top.yumbo.excel.annotation.ExcelCellStyle;
import top.yumbo.excel.annotation.ExcelTableHeader;
import top.yumbo.excel.entity.TitleBuilder;
import top.yumbo.excel.entity.TitleBuilders;
import top.yumbo.excel.util.BigDecimalUtils;

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
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author jinhua
 * @date 2021/6/22 10:34
 * <p>
 * 导出的设计思路：list->jsonArray->sheet
 * <p>
 * 可以以输入流方式传入待填入数据的模板excel文件
 * 1、默认样式导出
 * 2、高亮行导出
 */
public class ExcelExportUtils {

    //表头信息
    private enum TableEnum {
        WORK_BOOK, SHEET, TABLE_NAME, TABLE_HEADER, TABLE_HEADER_HEIGHT, RESOURCE, TYPE, TABLE_BODY, PASSWORD
    }

    // 单元格信息
    private enum CellEnum {
        TITLE_NAME, FIELD_NAME, FIELD_TYPE, SIZE, PATTERN, NULLABLE, WIDTH, EXCEPTION, COL, ROW, SPLIT, PRIORITY, FORMAT, MAP_ENTRIES
    }

    //样式的属性名
    private enum CellStyleEnum {
        FONT_NAME, FONT_SIZE, BG_COLOR, TEXT_ALIGN, LOCKED, HIDDEN, BOLD,
        VERTICAL_ALIGN, WRAP_TEXT, STYLES,
        FORE_COLOR, ROTATION, FILL_PATTERN, AUTO_SHRINK, TOP, BOTTOM, LEFT, RIGHT
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
     * 获取表头的高度
     */
    private static Integer getTableHeight(JSONObject tableHeaderDesc) {
        return tableHeaderDesc.getInteger(TableEnum.TABLE_HEADER_HEIGHT.name());
    }

    /**
     * 从json中获取Excel表头部分数据
     */
    private static JSONObject getTableHeaderDescInfo(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(TableEnum.TABLE_HEADER.name());
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
     * @param list
     * @param titleBuilders
     * @param outputStream
     * @param <T>
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
        exportExcelRowHighLight(list, fileInputStream, outputStream, function);
    }

    /**
     * 生成简单表头,需要提前构建好表头
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
     * 默认的导出（模板excel从注解信息注入）
     *
     * @param list         数据
     * @param outputStream 输出流
     */
    public static <T> void exportExcel(List<T> list, OutputStream outputStream) throws Exception {
        exportExcelRowHighLight(list, null, outputStream, null);

    }

    /**
     * 默认的excel导出
     *
     * @param list         数据
     * @param inputStream  输入流
     * @param outputStream 返回的输出流
     */
    public static <T> void exportExcel(List<T> list, InputStream inputStream, OutputStream outputStream) throws Exception {
        exportExcelRowHighLight(list, inputStream, outputStream, null);
    }

    /**
     * 高亮导出（模板文件从注解中resource获取）
     *
     * @param list         数据
     * @param outputStream 输出流
     * @param function     功能型函数，返回颜色
     */
    public static <T> void exportExcelRowHighLight(List<T> list, OutputStream outputStream, Function<T, IndexedColors> function) throws Exception {
        exportExcelRowHighLight(list, null, outputStream, function);
    }

    /**
     * 高亮行的方式导出
     *
     * @param list         数据集合
     * @param inputStream  excel输入流
     * @param outputStream 导出文件的输出流
     * @param function     功能型函数，返回颜色
     */
    public static <T> void exportExcelRowHighLight(List<T> list, InputStream inputStream, OutputStream outputStream, Function<T, IndexedColors> function) throws Exception {
        if (list != null && list.size() > 0 && outputStream != null) {
            Sheet sheet = null;
            Workbook workbook;
            // 1、如果输入流不为null，则从输入流中得到workbook
            if (inputStream != null) {
                workbook = WorkbookFactory.create(inputStream);//可以读取xls格式或xlsx格式。
                sheet = workbook.getSheetAt(0);
                inputStream.close();
            }
            final JSONObject exportInfo = getExportInfo(list.get(0).getClass(), sheet);
            final JSONObject titleInfo = getExcelBodyDescInfo(exportInfo);
            workbook = getWorkBook(exportInfo);
            sheet = getSheet(exportInfo);
            final Integer height = getTableHeight(getTableHeaderDescInfo(exportInfo));
            subListFilledSheet(list, titleInfo, sheet, height, function);
            // 将数据填充进sheet中，如果function不为null则可以高亮行
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
        } else if (list == null || list.size() == 0) {
            throw new NullPointerException("list不能为null 或 list数据量不能为0");
        } else {
            throw new NullPointerException("输出流不能为空");
        }

    }

    /**
     * 获取一个样式，如果没有就创建一个并且进行缓存
     */
    private static CellStyle getCellStyle(Workbook workbook, short index) {
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
            return cellStyle;
        }
    }

    private final static ValidatorFactory vf = Validation.buildDefaultValidatorFactory();
    private final static Validator validator = vf.getValidator();

    /**
     * 将集合数据填入表格
     */
    private static <T> int subListFilledSheet(List<T> subList, JSONObject titleInfo, Sheet sheet, int start, Function<T, IndexedColors> function) throws RuntimeException {
        final int length = subList.size();

        short index = IndexedColors.WHITE.getIndex();// 默认白色背景
        // 每一个子列表同一个cellStyle
        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = getCellStyle(workbook, index);
        for (int i = 0; i < length; i++) {
            int rowNum = start + i;
            Row row = sheet.getRow(rowNum);
            if (sheet.getRow(rowNum) == null) {
                row = sheet.createRow(rowNum);// 创建一行
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
                cellStyle = getCellStyle(workbook, index);
            }

            final JSONObject json = JSONObject.parseObject(JSONObject.toJSONString(t));
            // 可以并行处理单元格
            jsonToOneRow(row, json, titleInfo, exception, cellStyle);

            if (exception.get() != null) {
                throw exception.get();
            }
        }

        return length;
    }

    /**
     * 将一个实体数据填充到row中
     */
    private static void jsonToOneRow(Row row, JSONObject json, JSONObject titleInfo, AtomicReference<RuntimeException> exception, CellStyle cellStyle) {
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
     * 获取导出的信息
     *
     * @param clazz 泛型
     */
    private static JSONObject getExportInfo(Class<?> clazz) throws Exception {
        return getExportInfo(clazz, null);
    }

    /**
     * @author jinhua
     * @date 2021/6/4 1:10
     */
    @Data
    @Builder
    public static class CellStyleBuilder {
        private String fontName;
        private Integer fontSize;
        /**
         * 完整的颜色定义见：{@link IndexedColors}
         * 红10，白9，黑8，粉14，绿17，蓝12，
         * 灰22，金51，
         */
        private Integer fontColor;
        private Boolean bold;
        private Boolean locked;
        private Boolean hidden;
        private HorizontalAlignment textAlign;
        /**
         * 设置背景色
         * 完整的颜色定义见：{@link IndexedColors}
         * 默认白色，红10，白9，黑8，粉14，绿17，蓝12，
         * 灰22，金51，
         */
        private Integer bgColor;
        /**
         * 设置前景色（默认白色9）
         * 完整的颜色定义见：{@link IndexedColors}
         * 红10，白9，黑8，粉14，绿17，蓝12，
         * 灰22，金51，
         */
        private Integer foregroundColor;
        private Integer rotation;
        private VerticalAlignment verticalAlignment;
        private FillPatternType fillPatternType;
        private BorderStyle top;
        private BorderStyle bottom;
        private BorderStyle left;
        private BorderStyle right;
        private Boolean wrapText;
        private Boolean autoShrink;


        /**
         * 默认返回新版本的样式
         */
        public CellStyle getCellStyle() {
            return getCellStyle("xlsx");
        }

        /**
         * 构建一个样式
         *
         * @param type 如果是xls则type="xls" 如果是xlsx则type="xlsx"
         */
        public CellStyle getCellStyle(String type) {
            Workbook workbook;
            if ("xls".equals(type)) {
                workbook = new HSSFWorkbook();
            } else {
                workbook = new HSSFWorkbook();
            }
            return getCellStyle(workbook);
        }

        public CellStyle getCellStyle(Workbook wb) {
            CellStyle cellStyle = wb.createCellStyle();
            Font font = wb.createFont();
            if (!StringUtils.hasText(fontName)) fontName = "微软雅黑";
            if (fontSize == null) fontSize = 11;
            if (fontColor == null) fontColor = 8;
            if (bgColor == null) bgColor = 9;
            if (bold == null) bold = false;
            if (rotation == null) rotation = 0;
            if (foregroundColor == null) foregroundColor = 9;
            if (locked == null) locked = false;
            if (hidden == null) hidden = false;
            if (wrapText == null) wrapText = false;
            if (autoShrink == null) autoShrink = false;
            if (textAlign == null) textAlign = HorizontalAlignment.CENTER;
            if (verticalAlignment == null) verticalAlignment = VerticalAlignment.CENTER;
            if (fillPatternType == null) fillPatternType = FillPatternType.SOLID_FOREGROUND;
            if (top == null) top = BorderStyle.THIN;
            if (bottom == null) bottom = BorderStyle.THIN;
            if (left == null) left = BorderStyle.THIN;
            if (right == null) right = BorderStyle.THIN;

            font.setFontName(fontName);// 字体
            font.setFontHeightInPoints(fontSize.shortValue());//设置字体大小
            font.setColor(fontColor.shortValue());
            font.setBold(bold);
            cellStyle.setFont(font);
            cellStyle.setLocked(locked);// 设置是否上锁，默认否
            cellStyle.setAlignment(textAlign);// 默认居中
            cellStyle.setRotation(rotation.shortValue());// 文字的旋转角度
            cellStyle.setVerticalAlignment(verticalAlignment);// 设置垂直方向的对齐
            cellStyle.setFillPattern(fillPatternType);// 设置填充前景色
            cellStyle.setBorderTop(top);// 设置边框类型，上
            cellStyle.setBorderBottom(bottom);// 下
            cellStyle.setBorderLeft(left);// 左
            cellStyle.setBorderRight(right);// 右
            cellStyle.setWrapText(wrapText);// 是否多行显示文本
            cellStyle.setShrinkToFit(autoShrink);// 如果文本太长，控制单元格是否应自动调整大小以缩小以适应
            cellStyle.setFillForegroundColor(foregroundColor.shortValue());// 设置前景色
            cellStyle.setFillBackgroundColor(bgColor.shortValue());// 设置背景色
            cellStyle.setHidden(hidden);//
            cellStyle.setFont(font);
            return cellStyle;
        }
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
     * 从json中获取Excel身体部分数据
     */
    private static JSONObject getExcelBodyDescInfo(JSONObject fulledExcelDescData) {
        return fulledExcelDescData.getJSONObject(TableEnum.TABLE_BODY.name());
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
                String titleName = getStringCellValue(cell, String.class.getTypeName());
                // 如果标题相同找到了这单元格，获取单元格下标存入
                if (title.equals(titleName)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return index;
    }

    /**
     * 根据输入流返回workbook
     *
     * @param inputStream excel的输入流
     */
    private static Workbook getWorkBookByInputStream(InputStream inputStream) throws Exception {
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
    private static Sheet getSheetByInputStream(InputStream inputStream, int idx) throws Exception {
        final Workbook workbook = getWorkBookByInputStream(inputStream);
        final Sheet sheet = workbook.getSheetAt(idx);
        if (sheet == null) {
            throw new NullPointerException("序号为" + idx + "的sheet不存在");
        }
        return sheet;
    }

    /**
     * 获取excel工作簿
     *
     * @param resourcePath 资源路径
     */
    private static Workbook getWorkBookByResource(String resourcePath, String type) throws Exception {
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
                    // 是相对路径，springboot环境下，打成jar也有效
                    ClassPathResource classPathResource = new ClassPathResource(split[1]);
                    inputStream = classPathResource.getInputStream();
                }
                return getWorkBookByInputStream(inputStream);
            }
            throw new IllegalArgumentException("请带上协议头例如http://");
        } else {
            throw new IllegalArgumentException("资源地址不正确，配置的资源地址：" + resourcePath);
        }
    }
}
