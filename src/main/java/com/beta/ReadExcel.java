package com.beta;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;

/**
 * 将excel中的文件数据读取出来
 *
 * @author itaiit
 */
public class ReadExcel {

    public static void main(String[] args) throws IOException {
        /*
         * 1、首先创建Workbook
         * 2、获得sheet
         * 3、获得row
         * 4、获得cell
         */

        InputStream ein = new FileInputStream(ReadExcel.class.getResource("/").getPath() + "source.xls");
        Workbook wb = new HSSFWorkbook(ein);
        // 获得第几个sheet，excel文件默认有3个sheet
        Sheet sheet = wb.getSheetAt(0);
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        // 获得当前sheet的最后一行的索引
        int rownum = sheet.getLastRowNum();

        int cellnum;
        Row row = null;
        Cell cell = null;
        for (int i = 0; i <= rownum; i++) {
            // 得到第i行
            row = sheet.getRow(i);
            // 暂时不考虑空行
            // if (row == null){
            // System.out.println("-----row null-----");
            // continue;
            // }
            // 得到总列数
            cellnum = row.getLastCellNum();
            for (int j = 0; j < cellnum; j++) {
                // 如果excel中存在空的Cell则可以使用此句，将空的Cell转换为BLANK，进一步在switch中判断
                cell = row.getCell(j);
                if (cell == null) {
                    System.out.print(cell + "\t");
                    continue;
                } else {
                    // 得到该列的CellType类型
                    CellType ct = cell.getCellTypeEnum();
                    switch (ct) {
                        case STRING:
                            System.out.print(cell.getRichStringCellValue()
                                    .toString() + "\t");
                            break;
                        case NUMERIC:
                            // 如果数据格式符合时间日期的格式，则该数据为日期数据，将其看作日期，否则是数值
                            if (DateUtil.isCellDateFormatted(cell)) {
                                System.out
                                        .print(sdf.format(cell.getDateCellValue())
                                                + "\t");
                            } else {
                                System.out.print(cell.getNumericCellValue() + "\t");
                            }
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        default:
                            System.out.println();
                    }
                }
            }
            System.out.println();
        }
        ein.close();
    }
}
