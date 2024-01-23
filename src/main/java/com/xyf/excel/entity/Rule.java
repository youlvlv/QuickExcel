package com.xyf.excel.entity;

/**
 * 导入时校验规则
 *
 * @author lizhiwei
 */
public interface Rule<T> {

	void rule(T row);
}
