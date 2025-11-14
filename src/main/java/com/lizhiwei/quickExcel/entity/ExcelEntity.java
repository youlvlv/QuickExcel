package com.lizhiwei.quickExcel.entity;


import com.lizhiwei.quickExcel.format.ExcelFormat;
import com.lizhiwei.quickExcel.format.ExcelFormatBase;

import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Optional;

/**
 * 导出Excel实体类
 */
public class ExcelEntity {

	private static final String ALIAS_DEFAULT = "";

	/**
	 * 名称
	 */
	private String title;
	/**
	 * 别名
	 */
	private String alias = ALIAS_DEFAULT;
	/**
	 * 值
	 */
	private Integer value;
	/**
	 * 属性名
	 */
	private String property;
	/**
	 * 类型
	 */
	private Class<?> type;
	/**
	 * 转换器
	 */
	private ExcelFormatBase<?> format;
	/**
	 * 排序
	 */
	private int index;
	/**
	 * 顶部名称
	 */
	private String topName;
	/**
	 * 字段类型
	 */
	private ParamType paramType = ParamType.FIELD;
	/**
	 * 是否可读
	 */
	private boolean isRead = true;
	/**
	 * 是否可写
	 */
	private boolean isWrite = true;
	/**
	 * 读取时是否允许为空
	 */
	private boolean isNotNull = false;
	/**
	 * 属性别名
	 */
	private String aliasProperty = "";

	/**
	 * 精度
	 */
	private int accuracy = -1;
	/**
	 * 宽度
	 */
	private int width = 256 * 15;

    private Rule<?>[] rules;

	public boolean isRead() {
		return isRead;
	}

	public void setRead(boolean read) {
		isRead = read;
	}

	public boolean isWrite() {
		return isWrite;
	}

	public void setWrite(boolean write) {
		isWrite = write;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public Integer getValue() {
		return value;
	}

	public void setValue(Integer value) {
		this.value = value;
	}

	public String getProperty() {
		return property;
	}

	public void setProperty(String property) {
		this.property = property;
	}

	public Class<?> getType() {
		return type;
	}

	public void setType(Class<?> type) {
		this.type = type;
	}

	public ExcelFormatBase<?> getFormat() {
		return format;
	}

	public void setFormat(ExcelFormatBase<?> format) {
		this.format = format;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public String getTopName() {
		return topName;
	}

	public void setTopName(String topName) {
		this.topName = topName;
	}

	public ParamType getParamType() {
		return paramType;
	}

	public void setParamType(ParamType paramType) {
		this.paramType = paramType;
	}

	public TopName getTopNameInt(Class<? extends TopName> topName) {
		try {
			return topName.getDeclaredConstructor().newInstance();
		} catch (InstantiationException | NoSuchMethodException | InvocationTargetException |
		         IllegalAccessException e) {
			throw new RuntimeException(e);
		}
	}

	public boolean isNotNull() {
		return isNotNull;
	}

	public void setNotNull(boolean notNull) {
		isNotNull = notNull;
	}

	public int getWidth() {
		return width;
	}

	public void setWidth(int width) {
		this.width = width;
	}

	public String getAlias() {
		return alias;
	}

	public void setAlias(String alias) {
		this.alias = alias;
	}

	public int getAccuracy() {
		return accuracy;
	}

	public void setAccuracy(int accuracy) {
		this.accuracy = accuracy;
	}

	public String getEntityType(){
		return "normal";
	}

    public ExcelEntity(Integer value, String title, ExcelFormat<?> format) {
		this.title = title;
		this.value = value;
		this.format = format;
	}

	public ExcelEntity(Excel e, ExcelFormatBase<?> format, String value, Class clazz) {
        this.title = e.value();
        this.width = e.width();
        this.property = value;
        this.format = format;
        this.index = e.index();
        this.topName = e.topName();
        this.aliasProperty = e.aliasProperty();
        if (e.secondName() != DefaultTopName.class) {
            this.topName = getTopNameInt(e.secondName()).value();
        }
        if (!e.isPicture()) {
            this.paramType = e.type();
        } else {
            this.paramType = ParamType.IMAGE;
        }
        // 1. 获取 Class 数组
        Class<? extends Rule<?>>[] ruleClasses = e.rules();

        // 2. 转换为 Rule 实例数组
        Rule<?>[] rules = new Rule[ruleClasses.length];
        for (int i = 0; i < ruleClasses.length; i++) {
            try {
                // 假设 Rule 有无参构造函数
                rules[i] = ruleClasses[i].getDeclaredConstructor().newInstance();
            } catch (InstantiationException | IllegalAccessException |
                     InvocationTargetException | NoSuchMethodException ex) {
                throw new RuntimeException("Failed to instantiate rule: " + ruleClasses[i], ex);
            }
        }
        this.rules = rules;
        this.isRead = e.isRead();
        this.isWrite = e.isWrite();
        this.type = clazz;
        this.alias = e.alias();
        this.isNotNull = e.isNotNull();
        this.accuracy = e.accuracy();
	}

	public ExcelEntity(String value, String title, ExcelFormat<?> format, int index, Class<? extends TopName> topName, ParamType type) {
		this.title = title;
		this.property = value;
		this.format = format;
		this.index = index;
		this.topName = getTopNameInt(topName).value();
		this.paramType = type;
	}


	public ExcelEntity(String value, String title, ExcelFormat<?> format, int index, Class<? extends TopName> topName, Class clazz, ParamType type, boolean isRead, boolean isWrite, boolean isNotNull) {
		this.title = title;
		this.property = value;
		this.format = format;
		this.index = index;
		this.topName = getTopNameInt(topName).value();
		this.paramType = type;
		this.isRead = isRead;
		this.isWrite = isWrite;
		this.type = clazz;
		this.isNotNull = isNotNull;
	}

    public List<Rule<?>> getRules() {
        return List.of(Optional.of(rules).orElse(new Rule<?>[0]));
    }

	public ExcelEntity(ParamType index) {
		if (index == ParamType.INDEX) {
			this.title = "序号";
			this.property = "";
			this.paramType = index;
			this.type = Integer.class;
			this.topName = "";
			this.width = 256 * 15;
		}
	}

	public ExcelEntity(Integer value, String title) {
		this.title = title;
		this.value = value;
	}

	public ExcelEntity(String value, String title) {
		this.title = title;
		this.property = value;
	}

	public ExcelEntity() {
	}

    public String getAliasProperty() {
        return aliasProperty;
    }

    public void setAliasProperty(String aliasProperty) {
        this.aliasProperty = aliasProperty;
    }
}
