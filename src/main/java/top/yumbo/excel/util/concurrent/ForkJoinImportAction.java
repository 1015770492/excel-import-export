package top.yumbo.excel.util.concurrent;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import top.yumbo.excel.util.CheckLogicUtils;
import top.yumbo.excel.util.constants.CellEnum;

import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.RecursiveAction;
import java.util.function.Consumer;

import static top.yumbo.excel.util.ExcelImportExportUtils.*;
import static top.yumbo.excel.util.concurrent.ForkJoinImportTask.listToJoinString;
import static top.yumbo.excel.util.concurrent.ForkJoinImportTask.printRowOfException;

/**
 * @author jinhua
 * @date 2022/7/27 23:24
 */

public class ForkJoinImportAction<T> extends RecursiveAction {

    private final int start;
    private final int end;
    private final boolean recordAllException;
    private final Sheet sheet;
    private final int threshold;// 默认1万以后需要拆分
    private final int limitRowException;// 默认1万以后需要拆分
    private final JSONObject fieldInfo;// tableBody
    private final Class<T> clazz; // 泛型
    private final ValidatorFactory vf = Validation.buildDefaultValidatorFactory();
    private final Validator validator = vf.getValidator();
    private final Consumer<List<T>> consumer;

    public ForkJoinImportAction(JSONObject fieldInfo, Class<T> clazz, Consumer<List<T>> consumer,Sheet sheet, int start, int end, boolean recordAllException, int limitRowException, int threshold) {
        this.fieldInfo = fieldInfo;
        this.clazz = clazz;
        this.sheet = sheet;
        this.start = start;
        this.end = end;
        this.threshold = threshold;
        this.limitRowException = limitRowException;
        this.recordAllException = recordAllException;
        this.consumer = consumer;
    }


    @Override
    protected void compute() {
        int nums = end - start;// 计算有多少行数据

        if (nums <= threshold) {
            // 解析数据并且返回List
            List<T> list = praseRowsToList();
            // 消费对象
            consumer.accept(list);
        } else {
            int middle = (start + end) / 2;

            // 处理start到middle行号内的数据
            ForkJoinImportAction<T> left = new ForkJoinImportAction<>(fieldInfo, clazz, consumer,sheet, start, middle, recordAllException, limitRowException, threshold);
            left.fork();
            // 处理middle+1到end行号内的数据
            ForkJoinImportAction<T> right = new ForkJoinImportAction<>(fieldInfo, clazz,consumer, sheet, middle + 1, end, recordAllException, limitRowException, threshold);
            right.fork();
            left.join();
            right.join();
        }
    }

    /**
     * 解析从start到end行的数据转换为List
     */
    public List<T> praseRowsToList() throws RuntimeException {
        // 从表头描述信息得到表头的高
        final ArrayList<T> result = new ArrayList<>();

        ArrayList<List<String>> rowOfErrMessage = new ArrayList<>();
        // 按行扫描excel表
        for (int i = start; i <= end; i++) {
            final Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            JSONObject convertedRowJSONObject = new JSONObject();// 一行数据
            convertedRowJSONObject.put(CellEnum.ROW.name(), i + 1);// 记录行号
            // 错误消息列表
            ArrayList<String> errMessageList = new ArrayList<>();
            // 记录null的次数
            int countNull = 0;
            // 将row转JSON，如果有错误则计入到
            rowToJSONObjectWithRecordError(row, convertedRowJSONObject, errMessageList);

            // 判断这行数据是否正常
            // 正常情况下count是等于length的，因为每个字段都需要处理
            if (errMessageList.size() == 0) {
                if (countNull != convertedRowJSONObject.size() - 1) {
                    // 进行jsr303校验
                    try {
                        T t = JSONObject.parseObject(convertedRowJSONObject.toJSONString(), clazz);
                        t = CheckLogicUtils.checkNullLogicWithJSR303(t, validator);
                        result.add(t);// 正常情况下添加一条数据
                    } catch (RuntimeException e) {
                        errMessageList.add("第" + convertedRowJSONObject.getBigInteger(CellEnum.ROW.name()) + "行数据异常：" + e.getMessage());
                    }
                }

            } else {
                // 该行存在error
                rowOfErrMessage.add(errMessageList);
            }
            // 如果没有开启记录所有日志，则该task不会跑完，当达到了100条日志的时候就会抛异常
            if (!recordAllException) {
                if (rowOfErrMessage.size() > limitRowException) {
                    // 每个线程默认达到100Exception时就结束
                    throw new RuntimeException(Thread.currentThread().getName() + "-->超过" + limitRowException + "条异常记录\n" + printRowOfException(rowOfErrMessage));
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
        if (rowOfErrMessage.size() > 0) {
            System.err.println(printRowOfException(rowOfErrMessage));
        }

        return result;
    }

    /**
     * 将一行excel数据转化为JSONObject（相当于是中间对象）
     *
     * @param row            行
     * @param oneRow         转换后的对象
     * @param errMessageList 异常收集
     */
    private void rowToJSONObjectWithRecordError(Row row, JSONObject oneRow, ArrayList<String> errMessageList) {
        // 将Row转换为JSONObject
        fieldInfo.forEach((fieldName, arr) -> {
            // 字段的信息可能是一个数组，因为存在重复注解的情况
            JSONArray fieldDescArr = (JSONArray) arr;
            String join = null;
            int replaceType = 0;
            JSONObject temp = null;
            // 单个注解对字段的转换逻辑
            ArrayList<Object> tempList = new ArrayList<>();

            // 遍历注解数组，根据注解信息从 row 中读取对应的单元格（可能会读取多个单元格）
            for (Object obj : fieldDescArr) {
                JSONObject fieldDesc = (JSONObject) obj;
                //
                replaceType = fieldDesc.getInteger(CellEnum.REPLACE_ALL_TYPE.name());
                if (fieldDescArr.size() > 1) {
                    temp = fieldDesc;
                }
                try {
                    // 根据注解信息获取字段的值（如果字段只有一个注解，则可以进行类型转换）
                    Object value = getValue(row, fieldDesc, replaceType);
                    if (value != null) {
                        tempList.add(value);
                        if (join == null) {
                            join = fieldDesc.getString(CellEnum.JOIN.name());// 进行正则切割
                        } else if (!fieldDesc.getString(CellEnum.JOIN.name()).equals("$$")) {
                            join = fieldDesc.getString(CellEnum.JOIN.name());
                        }
                    }
                } catch (Exception e) {
                    errMessageList.add(e.getMessage());
                }
            }
            if (tempList.size() == 1) {
                // 情形1、说明是单个注解，那么可以根据javaBean定义的类型进行转换
                oneRow.put(fieldName, tempList.get(0));// 添加数据
            } else if (tempList.size() > 1) {
                // 情形2、多个注解的情况一定是字符串类型的，不然存不下2个不同注解的内容
                String mergedStr = listToJoinString(join, tempList);
                if (replaceType == 1) {
                    mergedStr = replaceAllOrReplacePart(mergedStr, temp);
                }
                oneRow.put(fieldName, mergedStr);
            }
        });
    }

}
