package com.study.excel.style.border;

import org.apache.poi.ss.usermodel.CellStyle;

public enum ExcelBorderValues {
	
	NONE(CellStyle.BORDER_NONE),
	THIN(CellStyle.BORDER_THIN),
	MEDIUM(CellStyle.BORDER_MEDIUM),
	DASHED(CellStyle.BORDER_DASHED),
	DOTTED(CellStyle.BORDER_DOTTED),
	THICK(CellStyle.BORDER_THICK),
	DOUBLE(CellStyle.BORDER_DOUBLE),
	HAIR(CellStyle.BORDER_HAIR),
	MEDIUM_DASHED(CellStyle.BORDER_MEDIUM_DASHED),
	DASH_DOT(CellStyle.BORDER_DASH_DOT),
	MEDIUM_DASH_DOT(CellStyle.BORDER_MEDIUM_DASH_DOT),
	DASH_DOT_DOT(CellStyle.BORDER_DASH_DOT_DOT),
	MEDIUM_DASH_DOT_DOT(CellStyle.BORDER_DASH_DOT_DOT),
	SLANTED_DASH_DOT(CellStyle.BORDER_SLANTED_DASH_DOT);

	private final short borderStyle;

	ExcelBorderValues(short borderStyle) {
		this.borderStyle = borderStyle;
	}

	public short getBorderStyle() {
		return borderStyle;
	}
}
