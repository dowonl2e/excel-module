package com.study.excel.style.background;

import org.apache.poi.ss.usermodel.CellStyle;

public class NoExcelColor implements ExcelBackgroundColor {

	@Override
	public void applyForegroundWithRGB(CellStyle cellStyle) {
		/* Do Nothing */
	}

	@Override
	public void applyBackgroundWithRGB(CellStyle cellStyle) {
		/* Do Nothing */
	}

}
