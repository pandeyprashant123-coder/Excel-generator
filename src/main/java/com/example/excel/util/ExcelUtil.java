package com.example.excel.util;

public class ExcelUtil {

    public static String safeName(String name) {
        return name.replaceAll("\\W+", "_");
    }

    public static String col(int index) {
        StringBuilder sb = new StringBuilder();
        while (index >= 0) {
            sb.insert(0, (char) ('A' + index % 26));
            index = index / 26 - 1;
        }
        return sb.toString();
    }
}
