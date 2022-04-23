package com.study.excel.style.align;

import org.apache.poi.ss.usermodel.CellStyle;

public enum ExcelAlignStyle implements ExcelAlign {

	LEFT_TOP(CellStyle.ALIGN_LEFT, CellStyle.VERTICAL_TOP),
	LEFT_CENTER(CellStyle.ALIGN_LEFT, CellStyle.VERTICAL_CENTER),
	LEFT_BOTTOM(CellStyle.ALIGN_LEFT, CellStyle.VERTICAL_BOTTOM),
	CENTER_TOP(CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_TOP),
	CENTER_CENTER(CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_CENTER),
	CENTER_BOTTOM(CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_BOTTOM),
	RIGHT_TOP(CellStyle.ALIGN_RIGHT, CellStyle.VERTICAL_TOP),
	RIGHT_CENTER(CellStyle.ALIGN_RIGHT, CellStyle.VERTICAL_CENTER),
	RIGHT_BOTTOM(CellStyle.ALIGN_RIGHT, CellStyle.VERTICAL_BOTTOM);
	
	private final short horizontal;
	private final short vertical;
	
	ExcelAlignStyle(short horizontal, short vertical) {
		this.horizontal = horizontal;
		this.vertical = vertical;
	}
	
	public short getHorizontal() {
		return horizontal;
	}

	public short getVertical() {
		return vertical;
	}
	
	@Override
	public void apply(CellStyle cellStyle) {
		cellStyle.setAlignment(getHorizontal());
		cellStyle.setVerticalAlignment(getVertical());
	}

}
