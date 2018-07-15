package com.example.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @name: com.example.poi.annotation.ExcelModel.java
 * @description:
 * @author: Wang Liang
 * @create Time: 2018/7/14 17:09
 * @copuright: 深圳拓保软件有限公司
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelModel {
	/**
	 * 导出标题名
	 */
	String chineseName();

	/**
	 * 导入属性索引
	 */
	int index() default -1;

	/**
	 * 是否状态值
	 */
	boolean isStatusValue() default false;
}
