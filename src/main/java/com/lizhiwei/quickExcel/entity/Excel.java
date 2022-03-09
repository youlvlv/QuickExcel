package com.lizhiwei.quickExcel.entity;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Excel {
    /**
     * 导入EXCEL匹配名称
     * @return
     */
    String value() default "";
    /**
     * 导出EXCEL时 匹配名称
     *
     * @return
     */
    String name() default "";

    /**
     * 默认排序
     * @return
     */
    int index() default -1;


    Class<? extends ExcelFormat> format() default DefaultFormat.class;
}
