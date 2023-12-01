package com.lizhiwei.quickExcel.entity;

import java.util.List;

public class ExcelEntityGroup extends ExcelEntity{

	private List<ExcelEntity> propertity;

	private String firstName;


	public List<ExcelEntity> getPropertity() {
		return propertity;
	}

	public void setPropertity(List<ExcelEntity> propertity) {
		this.propertity = propertity;
	}

	public String getFirstName() {
		return firstName;
	}

	public void setFirstName(String firstName) {
		this.firstName = firstName;
	}

	@Override
	public String getEntityType() {
		return "group";
	}
}
