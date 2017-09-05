package com.goodSoft.util;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by ASUS on 2017/7/7.
 */
@SuppressWarnings("ALL")
public class ReadExcel {
    // 默认单元格格式化日期字符串
    private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    // 格式化数字
    private static DecimalFormat nf = new DecimalFormat("0");

    public static List<List<Object>> readExcel(File file) {
        List<List<Object>> rowList = new ArrayList<>();
        List<Object> cellList = null;
        XSSFWorkbook wb = null;
        XSSFRow row;
        XSSFCell cell;
        try {
            wb = new XSSFWorkbook(new FileInputStream(file));
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
        int sheetNum = wb.getNumberOfSheets();
        //遍历sheet表
        for (int i = 0; i < sheetNum; ++i) {
            XSSFSheet sheet = wb.getSheetAt(i);
            //遍历行
            for (int j = sheet.getFirstRowNum() + 1, rowLength = sheet.getLastRowNum(); j <= rowLength; ++j) {
                row = sheet.getRow(j);
                cellList = new ArrayList<>();
                if (row == null && j != rowLength) {
                    rowList.add(cellList);
                }
                //遍历列
                for (int k = row.getFirstCellNum(), cellLength = row.getLastCellNum(); k < cellLength; ++k) {
                    cell = row.getCell(k);
                    if (cell == null && k != cellLength && cell.getCellType() == XSSFCell.CELL_TYPE_BLANK) {
                        break;
                    } else {
                        switch (cell.getCellType()) {
                            case XSSFCell.CELL_TYPE_STRING:
                                cellList.add(cell.getStringCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_FORMULA:
                                cellList.add(cell.getNumericCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                                    cellList.add(nf.format(cell.getNumericCellValue()));
                                } else if ("General".equals(cell.getCellStyle().getDataFormatString())) {
                                    cellList.add(nf.format(cell.getNumericCellValue()));
                                } else {
                                    cellList.add(sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue())));
                                }
                                break;
                            case XSSFCell.CELL_TYPE_BOOLEAN:
                                cellList.add(cell.getBooleanCellValue());
                                break;
                            default:
                                ++k;
                                cellList.add(null);
                        }
                    }
                }
                rowList.add(cellList);
            }
        }
        return rowList;
    }
}
