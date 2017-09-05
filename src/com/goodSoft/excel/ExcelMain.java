package com.goodSoft.excel;

import com.goodSoft.util.ReadExcel;

import java.io.File;
import java.util.List;

/**
 * Created by ASUS on 2017/7/7.
 */
public class ExcelMain {
    public static void main(String[] arg) {
        String fileName = "C:/Users/ASUS/Desktop/第二批档案邮寄地址.xlsx";
        File file = new File(fileName);
        List<List<Object>> list = ReadExcel.readExcel(file);
        for (int i = 0, length = list.size(); i < length; ++i) {
            List<Object> data = list.get(i);
            for (Object msg : data) {
                System.out.print(msg + " | ");
            }
            System.out.println();
            System.out.println("---------------------------------------------------------------------------------------");
        }
    }
}
