package top.yumbo.excel.util.concurrent;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.*;
import org.springframework.util.StringUtils;
import top.yumbo.excel.util.constants.CellEnum;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.RecursiveTask;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;
import java.util.function.Function;

import static top.yumbo.excel.util.ExcelImportExportUtils.getValueByFieldInfo;
import static top.yumbo.excel.util.ExcelImportExportUtils.setCellValue;

public class ForkJoinExportTask<T> extends RecursiveTask<Integer> {

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
        public ForkJoinExportTask(List<T> subList, JSONObject titleInfo, Sheet sheet, int start, int end, int threshold, Function<T, IndexedColors> function) {
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

