package com.lizhiwei.quickExcel.entity;

import com.lizhiwei.quickExcel.exception.ExcelValueException;

/**
 * 导入时校验规则
 *
 * @author lizhiwei
 */
public interface Rule<T> {

    void rule(T t) throws ExcelValueException;
}
