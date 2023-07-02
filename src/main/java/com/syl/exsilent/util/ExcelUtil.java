package com.syl.exsilent.util;

import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Method;

/**
 * 工具类
 *
 * @since 2023.6.23
 */
@Slf4j
public class ExcelUtil {
    /**
     * 获取对象字段值
     *
     * @param object    对象实例
     * @param fieldName 字段名称
     * @param <T>       泛型类
     * @return 对象实例的字段值
     */
    public static <T> Object getFieldVal(T object, String fieldName) {
        Object val = null;
        try {
            Class<?> tClass = object.getClass();
            Method method = tClass.getDeclaredMethod(getGetter(fieldName));
            val = method.invoke(object);
        } catch (Exception e) {
            log.error("reflect get field value failed", e);
        }
        return val;
    }

    /**
     * 首字母大写
     *
     * @param str 字符串
     * @return 首字母大写的字符串
     */
    public static String capitalizeFirst(String str) {
        return str.substring(0, 1).toUpperCase() + str.substring(1);
    }

    /**
     * 获取字段的get方法名称
     *
     * @param fieldName 方法名称
     * @return 字段的get方法名称
     */
    public static String getGetter(String fieldName) {
        return "get" + capitalizeFirst(fieldName);
    }

}
