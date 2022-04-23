package com.study.excel.data;

import java.util.ArrayList;
import java.util.List;

public class ExcelData<T> {

	private List<T> dataList = new ArrayList<T>();
	
	public ExcelData<T> addData(T value) {
		dataList.add(value);
		return this;
	}

	public void setDataList(List<T> dataList) {
		this.dataList = dataList;
	}

	public List<T> getDataList() {
		return this.dataList;
	}
	
	public ExcelData<T> clearData(){
		dataList = new ArrayList<>();
		return this;
	}
}
