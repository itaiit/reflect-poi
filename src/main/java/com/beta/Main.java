package com.beta;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * 实现将excel中的数据分装为javabean
 *
 * @author itaiit
 */
public class Main {

    /*
     * 1、加载excel文件，初始化参数
     */

    // 存放文件路径
    private String filepath = "source.xls";
    // 存储bean的相关信息
    private Class<?> cla = Emp.class;

    private List<Emp> emps = new ArrayList<Emp>();

    public List<Emp> read() throws IllegalArgumentException, IllegalAccessException, InstantiationException {

        InputStream fin;
        Workbook wb;
        Emp emp = null;

        try {
            fin = new FileInputStream(filepath);
            wb = new HSSFWorkbook(fin);
            Sheet sheet = wb.getSheetAt(0);

            // 获得当前sheet的最后一行的索引
            int rownum = sheet.getLastRowNum();

            int cellnum;
            Row row = null;
            Cell cell = null;

            // 存放第一行的列名
            row = sheet.getRow(0);
            // 创建一个数组存放列名
            String[] columnNames = new String[row.getLastCellNum()];
            for (int i = 0; i < columnNames.length; i++) {
                columnNames[i] = row.getCell(i).getRichStringCellValue().toString();
            }

            // 获得bean中的属性
            Field[] fields = cla.getDeclaredFields();

            // 遍历所有的行(不包含第0行)
            for (int i = 1; i <= rownum; i++) {
                // 必须在里面重新创建新的对象，否则对象都将指向同一块内存空间
                emp = (Emp) cla.newInstance();
                // 遍历属性名
                for (int k = 0; k < fields.length; k++) {
                    // 得到第i行
                    row = sheet.getRow(i);
                    // 暂时不考虑空行
                    if (row == null) {
                        continue;
                    }
                    // 得到总列数
                    cellnum = row.getLastCellNum();
                    // 遍历列
                    for (int j = 0; j < cellnum; j++) {
                        // 如果excel中存在空的Cell则可以使用此句，将空的Cell转换为BLANK，进一步在switch中判断
                        cell = row.getCell(j);

                        if (cell == null) {
                            continue;
                        }

                        /*
                         * 存放列名的数组的顺序与excel表中的数据顺序相同， 在赋值时，由于通过反射得到的fields数组中存放的
                         * field是顺序不固定的，因此在赋值时，需要循环遍历
                         *
                         * 当该索引下标所代表的列名等于列名数组中的列名时， 此时，该下表所对应的值即为该field的值。
                         */
                        if (columnNames[j].equalsIgnoreCase(fields[k].getName())) {
                            // 得到该列的CellType类型
                            CellType ct = cell.getCellTypeEnum();

                            // 分别取出每个属性
                            Field field = fields[k];
                            // 设置该属性可见
                            field.setAccessible(true);

                            switch (ct) {
                                case STRING:
                                    field.set(emp, cell.getRichStringCellValue().getString());
                                    break;
                                case NUMERIC:
                                    // 如果数据格式符合时间日期的格式，则该数据为日期数据，将其看作日期，否则是数值
                                    if (DateUtil.isCellDateFormatted(cell)) {
                                        field.set(emp, new java.sql.Date(cell.getDateCellValue().getTime()));
                                    } else {
                                        /*
                                         * javabean中定义的整数类型包含Integer和Double类型，
                                         * 需要根据类型做转换，后赋值
                                         */
                                        if ("java.lang.Integer".equals(field.getType().getCanonicalName())) {
                                            field.set(emp, new Double(cell.getNumericCellValue()).intValue());
                                        } else {
                                            field.set(emp, cell.getNumericCellValue());
                                        }
                                    }
                                    break;
                                case BOOLEAN:
                                    field.setBoolean(emp, cell.getBooleanCellValue());
                                    break;
                                default:
                                    System.out.println();
                            }
                        }
                    }
                }
                emps.add(emp);
                // 准备下一行的遍历
            }
            return emps;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return emps;
    }

    public static void main(String[] args) {
        /*
         * 已知与数据库中对应表的bean
         *
         * 已知存放数据的excel文件
         *
         * 需求：将excel中的数据封装到bean中，打印
         */

        try {
            Main tool = new Main();
            List<Emp> emps = tool.read();
            for (Emp emp : emps) {
                System.out.println(emp);
            }
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InstantiationException e) {
            e.printStackTrace();
        }
    }

}
