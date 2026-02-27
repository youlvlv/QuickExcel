package com.lizhiwei.quickExcel.model;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public interface ModelBase<T extends ModelBase> {

    static final Logger logger = LoggerFactory.getLogger(ModelBase.class);

    T over();
}
