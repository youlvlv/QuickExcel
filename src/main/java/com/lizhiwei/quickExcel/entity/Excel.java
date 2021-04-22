package com.lizhiwei.quickExcel.entity;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Excel {
    /**
     * 导出EXCEL时 匹配名称
     *
     * @return
     */
    String name() default "";

    String value() default "";

}
