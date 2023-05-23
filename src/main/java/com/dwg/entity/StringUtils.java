package com.dwg.entity;

/**
 * @author Autrui
 * @date 2023/5/23
 * @apiNote
 */
public class StringUtils {
    public static boolean isBlank(String value) {
        return null == value || "".equals(value) || "/".equals(value);
    }

}
