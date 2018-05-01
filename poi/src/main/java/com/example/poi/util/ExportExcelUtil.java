package com.example.poi.util;

import com.example.poi.annotation.ExportIgnore;
import com.example.poi.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Created by wangliang on 2017/8/31.
 */
public class ExportExcelUtil {

    /**
     * 指定数据导出到Excel
     *
     * @param data
     * @param out
     */
    public static <T> void exportExcel(List<T> data, String titleName, Map<String, String> titleMap, OutputStream out) {
        HSSFWorkbook workbook = null;
        try {
            titleName = getSheetName(data.get(0), titleName);
            // 创建工作博
            workbook = new HSSFWorkbook();
            // 合并单元格
            CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, 4);
            // 创建头标题样式
            HSSFCellStyle headStyle = createCellStyle(workbook, (short) 16);
            // 创建列标题样式
            HSSFCellStyle colStyle = createCellStyle(workbook, (short) 13);
            // sheet
            HSSFSheet sheet = workbook.createSheet(titleName);
            // 添加合并单元格对象
            sheet.addMergedRegion(cellRangeAddress);
            // 默认列宽度
            sheet.setDefaultColumnWidth(25);
            // 创建行
            HSSFRow row = sheet.createRow(0);
            // 创建单元格
            HSSFCell cell = row.createCell(0);
            // 加载单元格样式
            cell.setCellStyle(headStyle);
            cell.setCellValue(titleName);
            // 创建列标题
            HSSFRow titleRow = sheet.createRow(1);
            Field[] fields = getAllField(data.get(0));
            // 添加每列标题及样式
            for (int i = 0; i < fields.length; i++) {
                String fieldName = fields[i].getName();
                HSSFCell newCell = titleRow.createCell(i);
                newCell.setCellStyle(colStyle);
                newCell.setCellValue(titleMap.get(fieldName));
            }
            T t = null;
            // 创建单元格 写入数据
            if (data != null && !data.isEmpty()) {
                for (int i = 0; i < data.size(); i++) {
                    t = data.get(i);
                    Class<?> clazz = t.getClass();
                    // 创建数据行(前两行已经被占用)
                    HSSFRow dataRow = sheet.createRow(i + 2);
                    for (int j = 0; j < fields.length; j++) {
                        HSSFCell dataCell = dataRow.createCell(j);
                        String methodName = getterName(fields[j].getName());
                        Method method = clazz.getDeclaredMethod(methodName);
                        if (method == null) {
                            throw new IllegalArgumentException(clazz.getName() + " don't have method --> " + methodName);
                        }
                        Object result = method.invoke(t);
                        setCellValue(dataCell, String.valueOf(result));
                    }
               }
            }
            // 写入到文件
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // 关闭
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void setCellValue(HSSFCell dataCell, String value) {
        dataCell.setCellType(HSSFCell.CELL_TYPE_STRING);
        dataCell.setCellValue(value);
    }

    private static String getterName(String name) {
        return "get" + name.substring(0, 1).toUpperCase() + name.substring(1);
    }

    private static <T> String getSheetName(T t, String title) {
        if (title != null && !title.isEmpty()) {
            return title;
        }
        return t.getClass().getSimpleName();
    }

    private static <T> Field[] getAllField(T t) {
        List<Field> fields = new ArrayList<>();
        Class<?> clazz = t.getClass();
        Field[] declaredFields = clazz.getDeclaredFields();
        for (Field field : declaredFields) {
            ExportIgnore annotation = field.getAnnotation(ExportIgnore.class);
            if (annotation == null) {
                fields.add(field);
            }
        }
        return fields.toArray(new Field[fields.size()]);
    }


    /**
     * 单元格样式配置
     *
     * @param workbook
     * @param fontSize
     * @return
     */
    private static HSSFCellStyle createCellStyle(HSSFWorkbook workbook, short fontSize) {
        HSSFCellStyle style = workbook.createCellStyle();
        // 水平居中
        style.setAlignment(HorizontalAlignment.CENTER);
        // 垂直居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // 字体
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints(fontSize);
        // 加载字体
        style.setFont(font);
        return style;
    }

    /**
     * Excel文件导入
     *
     * @param file
     * @return
     */
    public static List<User> importExcel(File file) {
        FileInputStream inputStream = null;
        List<User> list = null;
        HSSFWorkbook workbook = null;
        try {
            list = new ArrayList<>();
            inputStream = new FileInputStream(file);
            // 读取文件
            workbook = new HSSFWorkbook(inputStream);
            // 读取sheet
            HSSFSheet sheet = workbook.getSheetAt(0);
            // 读取行(行数大于2)
            if (sheet.getPhysicalNumberOfRows() > 2) {
                User user = null;
                // 跳过前两行
                for (int i = 2; i < sheet.getPhysicalNumberOfRows(); i++) {
                    // 单元格
                    Row row0 = sheet.getRow(i);
                    user = new User();
                    // 封装数据
                    Cell cell0 = row0.getCell(0);
                    user.setName(cell0.getStringCellValue());
                    Cell cell1 = row0.getCell(1);
                    user.setAge(cell1.getStringCellValue());
                    Cell cell2 = row0.getCell(2);
                    user.setSex(cell2.getStringCellValue().equals("男") ? 1 : 0);
                    Cell cell3 = row0.getCell(3);
                    user.setEmail(cell3.getStringCellValue());
                    Cell cell4 = row0.getCell(4);
                    user.setPhone(cell4.getStringCellValue());
                    list.add(user);
                }
            }
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return list;
    }
}
