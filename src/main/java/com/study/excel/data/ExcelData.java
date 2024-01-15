package com.study.excel.data;

import org.springframework.lang.NonNull;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class ExcelData<T> {

	private List<T> dataList = new ArrayList<T>();
	
	public ExcelData<T> addData(@NonNull T value) {
		dataList.add(value);
		return this;
	}

	public ExcelData<T> addData(@NonNull T... values) {
		dataList = Arrays.stream(values).collect(Collectors.toList());
		return this;
	}

	public void setDataList(List<T> dataList) {
		this.dataList = dataList;
	}

	public List<T> getDataList() {
		return this.dataList;
	}

	/**
	 * 행 데이터 초기화
	 * @return ExcelData<?>
	 */
	public ExcelData<T> clearData(){
		dataList = new ArrayList<>();
		return this;
	}
}
