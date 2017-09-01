package com.example.poi.util;

import com.example.poi.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

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
    public static void exportExcel(List<User> data, OutputStream out) {
        HSSFWorkbook workbook = null;
        try {
            // 创建工作博
            workbook = new HSSFWorkbook();
            // 合并单元格
            CellRangeAddress cellRangeAddress = new CellRangeAddress(0, 0, 0, 4);
            // 创建头标题样式
            HSSFCellStyle headStyle = createCellStyle(workbook, (short) 16);
            // 创建列标题样式
            HSSFCellStyle colStyle = createCellStyle(workbook, (short) 13);
            // sheet
            HSSFSheet sheet = workbook.createSheet("用户名单");
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
            cell.setCellValue("用户列表");
            // 创建列标题
            HSSFRow titleRow = sheet.createRow(1);
            String[] titles = {"用户名", "年龄", "性别", "邮箱", "手机"};
            // 添加每列标题及样式
            for (int i = 0; i < titles.length; i++) {
                HSSFCell newCell = titleRow.createCell(i);
                newCell.setCellStyle(colStyle);
                newCell.setCellValue(titles[i]);
            }
            // 创建单元格 写入数据
            if (data != null) {
                for (int i = 0; i < data.size(); i++) {
                    User user = data.get(i);
                    // 写入每行数据(前两行已经被占用)
                    HSSFRow newRow = sheet.createRow(i + 2);
                    // 姓名
                    HSSFCell c1 = newRow.createCell(0);
                    c1.setCellValue(user.getName());
                    // 年龄
                    HSSFCell c2 = newRow.createCell(1);
                    c2.setCellValue(user.getAge());
                    // 性别
                    HSSFCell c3 = newRow.createCell(2);
                    c3.setCellValue(user.getSex() == 1 ? "男" : "女");
                    // 邮箱
                    HSSFCell c4 = newRow.createCell(3);
                    c4.setCellValue(user.getEmail());
                    // 手机
                    HSSFCell c5 = newRow.createCell(4);
                    c5.setCellValue(user.getPhone());
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
