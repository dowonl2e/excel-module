package com.study.excel.style.border;

import org.apache.poi.ss.usermodel.CellStyle;

public interface ExcelBorder {

	void applyTop(CellStyle cellStyle);
	void applyLeft(CellStyle cellStyle);
	void applyBottom(CellStyle cellStyle);
	void applyRight(CellStyle cellStyle);
	
	void apply(CellStyle cellStyle);
}
