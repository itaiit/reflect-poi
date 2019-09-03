package com.core;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class EtBUtil<T> {

    private Class<T> cla = null;
    // 存放要返回的封装对象的列表
    private List<T> objs = new ArrayList<T>();

    private List<String> colNames = new ArrayList<String>();

    // 存放javabean中的属性值
    private static Field[] fields = null;

    private EtBUtil(Class<T> cla) {
        this.cla = cla;

        fields = cla.getDeclaredFields();
    }

    /**
     * 读取文件中的记录，对记录进行处理
     *
     * @param filepath
     * @return
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public List<T> read(String filepath) throws IllegalAccessException, InstantiationException {

        if (filepath != null && !"".equalsIgnoreCase(filepath)) {
            try {
                Workbook wb = WorkbookFactory.create(new File(filepath));
                Sheet sheet = wb.getSheetAt(0);
                // 获得最后一行的记录的索引，从0开始
                int rownum = sheet.getLastRowNum();
                if (rownum <= 0) {
                    return null;
                }

                Row row0 = sheet.getRow(0);
                // 将第0行的字段名存放在集合中
                for (Cell cell : row0) {
                    colNames.add(cell.getRichStringCellValue().toString());
                }
                int cellNum = row0.getLastCellNum();

                T obj = null;

                for (int i = 1; i <= rownum; i++) {

                    Row rowi = sheet.getRow(i);

                    try {
                        obj = cla.newInstance();
                    } catch (InstantiationException e1) {
                        e1.printStackTrace();
                        throw e1;
                    }

                    // 双层循环，确保将excel中对应的值赋值给对应的属性，
                    // 由于通过反射得到的属性数组中，元素的存放顺序与声明顺序不一致
                    // 也就是说是无序的
                    for (int j = 0; j < cellNum; j++) {

                        Cell cellj = rowi.getCell(j);

                        if (cellj == null) {
                            continue;
                        }

                        for (Field field : fields) {
                            if (colNames.get(j).equals(field.getName())) {
                                // 设置javabean属性可见
                                field.setAccessible(true);

                                CellType ct = cellj.getCellTypeEnum();

                                switch (ct) {
                                    case STRING:
                                        try {
                                            field.set(obj, cellj.getRichStringCellValue().toString());
                                        } catch (IllegalArgumentException e) {
                                            e.printStackTrace();
                                            throw e;
                                        } catch (IllegalAccessException e) {
                                            e.printStackTrace();
                                            throw e;
                                        }
                                        break;
                                    case NUMERIC:
                                        /*
                                         * 如果是数值类型，那末判断：
                                         * 符合日期条件的类型，则为日期类型；
                                         * 否则，为基本数据类型，还需判断是Integer还是Double
                                         *
                                         */
                                        if (DateUtil.isCellDateFormatted(cellj)) {
                                            try {
                                                field.set(obj, new java.sql.Date(cellj.getDateCellValue().getTime()));
                                            } catch (IllegalArgumentException e) {
                                                e.printStackTrace();
                                                throw e;
                                            } catch (IllegalAccessException e) {
                                                e.printStackTrace();
                                                throw e;
                                            }
                                        } else {

                                            if ("java.lang.Integer".equals(field.getType().getCanonicalName())) {
                                                try {
                                                    field.set(obj, new Double(cellj.getNumericCellValue()).intValue());
                                                } catch (IllegalArgumentException e) {
                                                    e.printStackTrace();
                                                    throw e;
                                                } catch (IllegalAccessException e) {
                                                    e.printStackTrace();
                                                    throw e;
                                                }
                                            } else {
                                                try {
                                                    field.set(obj, cellj.getNumericCellValue());
                                                } catch (IllegalArgumentException e) {
                                                    e.printStackTrace();
                                                    throw e;
                                                } catch (IllegalAccessException e) {
                                                    e.printStackTrace();
                                                    throw e;
                                                }
                                            }
                                        }

                                        break;
                                    case BOOLEAN:
                                        try {
                                            field.setBoolean(obj, cellj.getBooleanCellValue());
                                        } catch (IllegalArgumentException e) {
                                            e.printStackTrace();
                                            throw e;
                                        } catch (IllegalAccessException e) {
                                            e.printStackTrace();
                                            throw e;
                                        }
                                        break;
                                    case BLANK:
                                    default:
                                        break;
                                }
                            } else {
                                continue;
                            }
                        }
                    }
                    objs.add(obj);
                }
                return objs;

            } catch (EncryptedDocumentException e) {
                e.printStackTrace();
            } catch (InvalidFormatException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    /**
     * 实例化一个工具类
     *
     * @param cla
     * @return
     */
    public static <T> EtBUtil<T> create(Class<T> cla) {
        return new EtBUtil<T>(cla);
    }
}
