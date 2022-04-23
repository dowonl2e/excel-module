package com.study.excel.style.background;

import org.apache.poi.ss.usermodel.CellStyle;

public interface ExcelBackgroundColor {

	void applyForegroundWithRGB(CellStyle cellStyle);

	void applyBackgroundWithRGB(CellStyle cellStyle);

}
