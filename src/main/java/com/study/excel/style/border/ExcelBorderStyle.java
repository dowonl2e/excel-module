package com.study.excel.style.border;

import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelBorderStyle implements ExcelBorder {
	
	private ExcelBorderValues borderValues;

	private ExcelBorderStyle(ExcelBorderValues borderValues) {
		this.borderValues = borderValues;
	}
	
	public static ExcelBorderStyle newInstance(ExcelBorderValues borderValues) {
		return new ExcelBorderStyle(borderValues);
	}
	
	@Override
	public void applyTop(CellStyle cellStyle) {
		cellStyle.setBorderTop(borderValues.getBorderStyle());
	}

	@Override
	public void applyLeft(CellStyle cellStyle) {
		cellStyle.setBorderLeft(borderValues.getBorderStyle());
	}

	@Override
	public void applyBottom(CellStyle cellStyle) {
		cellStyle.setBorderBottom(borderValues.getBorderStyle());
	}

	@Override
	public void applyRight(CellStyle cellStyle) {
		cellStyle.setBorderRight(borderValues.getBorderStyle());
	}

	@Override
	public void apply(CellStyle cellStyle) {
		applyTop(cellStyle);
		applyLeft(cellStyle);
		applyBottom(cellStyle);
		applyRight(cellStyle);
	}
	
}
