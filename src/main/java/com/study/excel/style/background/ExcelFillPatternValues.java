package com.study.excel.style.background;

import org.apache.poi.ss.usermodel.CellStyle;

public enum ExcelFillPatternValues {
	
	NONE(CellStyle.NO_FILL),
	SOLID(CellStyle.SOLID_FOREGROUND);
	
	private short pattern;

	private ExcelFillPatternValues(short pattern) {
		this.pattern = pattern;
	}

	public short getPattern() {
		return pattern;
	}

}
