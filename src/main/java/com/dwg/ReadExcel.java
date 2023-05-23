package com.dwg;

import com.dwg.service.ParseServiceExcelImpl;

import java.io.IOException;
import java.util.Scanner;

public class ReadExcel {

    public static void main(String[] args) {
        System.out.println("请输入excel文件路径: [execl文件路径] ");
        Scanner scanner = new Scanner(System.in);
        String[] strs = scanner.nextLine().split("\\s+");
        String filepath = strs[0];
        try {
            //将提供的execl文档中的点位编码保存到camera表中的crossing_number字段中
            ParseServiceExcelImpl.parseExcel(filepath);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
