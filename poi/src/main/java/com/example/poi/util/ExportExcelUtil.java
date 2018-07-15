package com.example.poi.util;

import com.example.poi.annotation.ExcelModel;
import com.example.poi.annotation.ImportModel;
import com.example.poi.entity.User;
import org.apache.commons.beanutils.BeanUtils;
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
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * Created by wangliang on 2017/8/31.
 *
 * @author wangliang
 */
public class ExportExcelUtil {

	public static final int ROWS = 2;

	public static <T> void export(List<T> data, List<String> titleNames, List<String> columns, OutputStream out) {
		HSSFWorkbook workbook = null;
		try {
			workbook = new HSSFWorkbook();
			// 标题样式
			HSSFCellStyle cellStyle = createCellStyle(workbook, (short) 13);
			HSSFCellStyle valueStyle = createValueStyle(workbook, (short) 10);
			// Sheet
			HSSFSheet sheet = workbook.createSheet();
			sheet.setDefaultColumnWidth(25);
			// 添加标题
			HSSFRow row = sheet.createRow(0);
			for (int i = 0; i < titleNames.size(); i++) {
				HSSFCell cell = row.createCell(i);
				cell.setCellStyle(cellStyle);
				cell.setCellValue(titleNames.get(i));
			}
			T t;
			// 数据写入
			for (int i = 0; i < data.size(); i++) {
				t = data.get(i);
				Class<?> target = t.getClass();
				// 创建数据行
				HSSFRow dataRow = sheet.createRow(i + 1);
				for (int j = 0; j < columns.size(); j++) {
					HSSFCell cell = dataRow.createCell(j);
					String getterName = getterName(columns.get(j));
					Method method = target.getDeclaredMethod(getterName);
					if (method != null) {
						Object result = method.invoke(t);
						cell.setCellStyle(valueStyle);
						cell.setCellValue(String.valueOf(result));
					}
				}
			}
			workbook.write(out);
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 指定数据导出到Excel
	 *
	 * @param data
	 * @param out
	 */
	public static <T> void exportExcel(List<T> data, String titleName, Map<String, String> titleMap, OutputStream out) {
		HSSFWorkbook workbook = null;
		try {
			if (titleName == null && titleName.isEmpty()) {
				titleName = data.get(0).getClass().getSimpleName();
			}
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
			List<Field> fields = getAllField(data.get(0).getClass());
			// 添加每列标题及样式
			for (int i = 0; i < fields.size(); i++) {
				String fieldName = fields.get(i).getName();
				HSSFCell newCell = titleRow.createCell(i);
				newCell.setCellStyle(colStyle);
				newCell.setCellValue(titleMap.get(fieldName));
			}
			T t;
			// 创建单元格 写入数据
			if (data != null && !data.isEmpty()) {
				for (int i = 0; i < data.size(); i++) {
					t = data.get(i);
					Class<?> clazz = t.getClass();
					// 创建数据行(前两行已经被占用)
					HSSFRow dataRow = sheet.createRow(i + ROWS);
					for (int j = 0; j < fields.size(); j++) {
						HSSFCell dataCell = dataRow.createCell(j);
						Field field = fields.get(j);
						if (field != null) {
							String methodName = getterName(field.getName());
							Method method = clazz.getDeclaredMethod(methodName);
							if (method == null) {
								throw new IllegalArgumentException(clazz.getName() + " don't have method --> " + methodName);
							}
							Object result = method.invoke(t);
							setCellValue(dataCell, String.valueOf(result));
						}
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
				if (workbook != null) {
					workbook.close();
				}
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

	/**
	 * 获取全部需要导入/导出的字段
	 *
	 * @param clazz
	 * @return
	 */
	private static List<Field> getAllField(Class<?> clazz) {
		List<Field> resultList = new ArrayList<>();
		Field[] fields = clazz.getDeclaredFields();
		for (Field field : fields) {
			if (field.isAnnotationPresent(ExcelModel.class)) {
				ExcelModel excelModel = field.getAnnotation(ExcelModel.class);
                if (excelModel.index() != -1) {
                    resultList.add(field);
                }
			}
		}
		return resultList;
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
	 * 单元格样式配置
	 *
	 * @param workbook
	 * @param fontSize
	 * @return
	 */
	private static HSSFCellStyle createValueStyle(HSSFWorkbook workbook, short fontSize) {
		HSSFCellStyle style = workbook.createCellStyle();
		// 水平居中
		style.setAlignment(HorizontalAlignment.CENTER);
		// 垂直居中
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		// 字体
		HSSFFont font = workbook.createFont();
		font.setFontHeightInPoints(fontSize);
		// 加载字体
		style.setFont(font);
		return style;
	}

	/**
	 * Excel Import To
	 *
	 * @param request
	 */
	public static List importExcel(File file, final Class<?> clazz, boolean haveTitle) {
		List resultList = new ArrayList<>();
		HSSFWorkbook workbook;
		try {
			List<Field> fields = getAllField(clazz);
//			CommonsMultipartResolver multipartResolver = new CommonsMultipartResolver(request.getSession().getServletContext());
//			if (multipartResolver.isMultipart(request)) {
//				MultipartRequest multipartRequest = (MultipartRequest) request;
//				Iterator<String> fileNames = multipartRequest.getFileNames();
//				while (fileNames.hasNext()) {
//					MultipartFile file = multipartRequest.getFile(fileNames.next());
//					if (file != null) {
						workbook = new HSSFWorkbook(new FileInputStream(file));
						HSSFSheet sheet = workbook.getSheetAt(0);
						// 数据行数
						int rows = sheet.getPhysicalNumberOfRows();
						if (haveTitle) {
							resultList.addAll(getCellValue(fields, rows, 1, sheet, clazz));
						} else {
							resultList.addAll(getCellValue(fields, rows, 0, sheet, clazz));
						}
//					}
//				}
//			}
		} catch (IOException e) {

		} catch (Exception e) {
			e.printStackTrace();
		}
		return resultList;
	}

	/**
	 * 得到数据
	 *
	 * @param fields		全部需要导入的字段
	 * @param startIndex	起始行索引
	 * @param rows
	 * @return
	 */
	private static List<ImportModel> getCellValue(List<Field> fields, int rows, int startIndex, HSSFSheet sheet, Class<?> clazz) throws Exception {
		List<ImportModel> resultList = new ArrayList<>();
		for (int i = startIndex; i < rows; i++) {
			ImportModel model = (ImportModel) clazz.newInstance();
			for (Field field : fields) {
				ExcelModel excelModel = field.getAnnotation(ExcelModel.class);
				HSSFCell cell = sheet.getRow(i).getCell(excelModel.index());
				Class<?> fieldType = field.getType();
				if (Date.class.equals(fieldType)) {
					BeanUtils.setProperty(model, field.getName(), cell.getDateCellValue().getTime());
				} else if (Integer.class.equals(fieldType)) {
					BeanUtils.setProperty(model, field.getName(), Integer.parseInt(cell.getStringCellValue()));
				} else if (BigDecimal.class.equals(fieldType)) {
					BeanUtils.setProperty(model, field.getName(), new BigDecimal(cell.getNumericCellValue()));
				} else {
					BeanUtils.setProperty(model, field.getName(), cell.getStringCellValue());
				}
			}
			resultList.add(model);
		}
		return resultList;
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
			if (sheet.getPhysicalNumberOfRows() > ROWS) {
				User user = null;
				// 跳过前两行
				for (int i = ROWS; i < sheet.getPhysicalNumberOfRows(); i++) {
					// 单元格
					Row row0 = sheet.getRow(i);
					user = new User();
					// 封装数据
					Cell cell0 = row0.getCell(0);
					user.setName(cell0.getStringCellValue());
					Cell cell1 = row0.getCell(1);
					user.setAge(cell1.getStringCellValue());
					Cell cell2 = row0.getCell(ROWS);
					user.setSex("男".equals(cell2.getStringCellValue()) ? 1 : 0);
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
