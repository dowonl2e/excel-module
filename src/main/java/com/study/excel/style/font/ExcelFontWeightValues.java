package com.study.excel.style.font;

import org.apache.poi.ss.usermodel.Font;

public enum ExcelFontWeightValues {
	
	NORMAL(Font.BOLDWEIGHT_NORMAL),
	BOLD(Font.BOLDWEIGHT_BOLD);
	
	private final short weight;

	private ExcelFontWeightValues(short weight) {
		this.weight = weight;
	}

	public short getWeight() {
		return weight;
	}

}
