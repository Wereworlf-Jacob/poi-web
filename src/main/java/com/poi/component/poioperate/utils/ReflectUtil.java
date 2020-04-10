package com.poi.component.poioperate.utils;

import com.poi.component.poioperate.vo.DbTableIncrement;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * @ClassName ReflectUtil
 * @Description 反射工具类
 * @Author Jacob
 * @Version 1.0
 * @since 2020/4/10 12:21
 **/
public class ReflectUtil {

    private static String GET = "get";

    private static String SET = "set";

    /**
     * 根据类的信息，获取该类中的set方法集合
     * @param clazz 类信息
     * @return {@link List< Method>} set方法列表集合
     * @throws
     *
     * @author Jacob
     * @since 2020/4/10
     */
    public static List<Method> getClassSetMethod(Class clazz) {
        return getClassMethod(clazz, SET);
    }

    /**
     * 根据类的信息，获取该类中get方法集合
     * @param clazz
     * @return {@link List< Method>}
     * @throws
     *
     * @author Jacob
     * @since 2020/4/10
     */
    public static List<Method> getClassGetMethod(Class clazz) {
        return getClassMethod(clazz, GET);
    }

    /**
     * 根据类型获取类的方法信息
     * @param clazz 类信息
     * @param type 方法类型 get/set
     * @return {@link List< Method>} 方法的集合
     * @throws
     *
     * @author Jacob
     * @since 2020/4/10
     */
    private static List<Method> getClassMethod(Class clazz, String type) {
        try {
            Map<String, Method> map = new LinkedHashMap<>();

            for (Field declaredField : clazz.getDeclaredFields()) {
                map.put(type + declaredField.getName().toLowerCase(), null);
            }
            for (Method method : clazz.getDeclaredMethods()) {
                String methodName = method.getName().toLowerCase();
                if (map.containsKey(methodName)) {
                    map.put(methodName, method);
                }
            }
            return map.entrySet().parallelStream().map(Map.Entry::getValue).
                    filter(a -> a != null).collect(Collectors.toList());
        } catch (Exception e) {
            throw new ClassCastException("class " + clazz.getName() + " has invalid properties it must all String type ");
        }
    }

}
