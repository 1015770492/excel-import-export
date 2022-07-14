package top.yumbo.excel.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import top.yumbo.excel.annotation.business.CheckNullLogic;
import top.yumbo.excel.annotation.business.CheckValues;
import top.yumbo.excel.annotation.business.MapEntry;

import javax.validation.ConstraintViolation;
import javax.validation.Validation;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import java.lang.reflect.Field;
import java.util.Set;

/**
 * @author jinhua
 * @date 2021/7/5 10:27
 * 校验逻辑工具类，内置jsr303校验
 */
public class CheckLogicUtils {

    /**
     * 校验空逻辑
     */
    public static <T> T checkNullLogicWithJSR303(T t) {
        Class<?> clazz = t.getClass();
        JSONObject fieldMap = getFieldMap(clazz);
        T entity = (T) checkNullLogic(JSONObject.parseObject(JSON.toJSONString(t)), clazz, fieldMap);
        // 校验值
        checkValue(t, clazz);
        ValidatorFactory vf = Validation.buildDefaultValidatorFactory();
        Validator validator = vf.getValidator();
        // 进行jsr303校验
        Set<ConstraintViolation<T>> set = validator.validate(entity);
        for (ConstraintViolation<T> constraintViolation : set) {
            throw new RuntimeException(constraintViolation.getMessage());
        }
        return entity;
    }

    /**
     * 强校验，用于校验字典
     */
    private static <T> void checkValue(T t, Class<?> clazz) {
        JSONObject data = JSONObject.parseObject(JSON.toJSONString(t));
        for (Field field : clazz.getDeclaredFields()) {
            CheckValues checkValues = field.getDeclaredAnnotation(CheckValues.class);
            if (checkValues != null) {
                int count = 0;
                for (String value : checkValues.values()) {
                    if (value != null && value.equals(data.get(field.getName()).toString())) {
                        // 通过
                        count = 1;
                        break;
                    }
                }
                if (count != 1) {
                    throw new RuntimeException(checkValues.message());
                }
            }
        }
    }

    /**
     * 校验非空逻辑
     */
    private static <T> T checkNullLogic(JSONObject data, Class<T> tClass, JSONObject fieldMap) {
        for (Field field : tClass.getDeclaredFields()) {
            CheckNullLogic annotation = field.getDeclaredAnnotation(CheckNullLogic.class);
            if (annotation != null) {
                // 校验follow的字段值是否符合values中的值
                String follow = annotation.follow();
                // 字典项的值
                String[] values = annotation.values();
                // 消息
                String message = annotation.message();
                // 校验标题的值
                String followTitle = annotation.followTitle();
                // 当前标题
                String title = annotation.title();
                boolean needCheck = false;
                for (String value : values) {
                    // 判断是否需要逻辑
                    if (org.springframework.util.StringUtils.hasText(follow)) {
                        // 获取字典的值
                        Object obj = data.get(follow);
                        // 1、看下是否需要较空，如果符合条件，需要进行较空
                        if (obj != null && value.equals(obj.toString())) {
                            // 需要进行校验，因为符合字典项
                            needCheck = true;
                            // 数据符合，接着校验当前field对应的值是否为null
                            Object fieldValue = data.get(field.getName());
                            // 2、校验当前字段值是否为值，
                            if (fieldValue == null || !org.springframework.util.StringUtils.hasText(fieldValue.toString())) {
                                // 2.1、等于空，需要抛异常。抛异常后就停下来了，不走后面的逻辑
                                if (org.springframework.util.StringUtils.hasText(title)) {
                                    // 找到对应的字典项值
                                    JSONObject titleMap = fieldMap.getJSONObject(follow);
                                    if (titleMap != null) {
                                        titleMap.forEach((k, v) -> {
                                            if (k.equals(value)) {
                                                throw new RuntimeException("'" + followTitle + "'为: '" + v + "' 时', " + title + "' 不能为空");
                                            }
                                        });
                                    }
                                } else {
                                    // 直接把消息抛出
                                    throw new RuntimeException(message);
                                }
                            } else {
                                // 2.2、不为null,说明校验通过了
                                break;
                            }
                        }
                    }

                }
                if (needCheck) {
                    // 校验空通过，处理下一个
                    continue;
                }
                // 不包含，本身这个数据不需要收集，故置null
                data.put(field.getName(), null);
            }
        }
        return JSONObject.parseObject(data.toJSONString(), tClass);
    }

    /**
     * 一次性的全部取出来
     */
    private static <T> JSONObject getFieldMap(Class<T> clazz) {
        JSONObject result = new JSONObject();
        for (Field field : clazz.getDeclaredFields()) {
            JSONObject fieldMap = new JSONObject();
            MapEntry[] mapEntries = field.getDeclaredAnnotationsByType(MapEntry.class);
            if (mapEntries != null && mapEntries.length > 0) {
                for (MapEntry mapEntry : mapEntries) {
                    String key = mapEntry.key();// key是中文
                    String value = mapEntry.value();// value是字典项
                    // 标题由message提供
                    fieldMap.put(value, key);
                }
                result.put(field.getName(), fieldMap);
            }

        }
        return result.entrySet().size() == 0 ? null : result;
    }

}
