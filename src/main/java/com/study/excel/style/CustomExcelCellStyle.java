package com.study.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import com.study.excel.configurer.ExcelCellStyleConfigurer;
import com.study.excel.style.align.ExcelAlign;
import com.study.excel.style.border.ExcelBorder;
import com.study.excel.style.font.ExcelFont;

public class CustomExcelCellStyle implements ExcelCellStyle {
	
	private ExcelCellStyleConfigurer configurer = new ExcelCellStyleConfigurer();
	
	public CustomExcelCellStyle() {}

	@Override
	public CustomExcelCellStyle align(ExcelAlign align) {
		configurer.align(align);
		return this;
	}
	
	@Override
	public CustomExcelCellStyle border(ExcelBorder border) {
		configurer.border(border);
		return this;
	}
	
	@Override
	public CustomExcelCellStyle foregroundColor(int red, int green, int blue) {
		configurer.foregroundColor(red, green, blue);
		return this;
	}
	
	@Override
	public CustomExcelCellStyle fontColor(int red, int green, int blue) {
		configurer.fontColor(red, blue, green);
		return this;
	}
	
	@Override
	public CustomExcelCellStyle fontWeight(ExcelFont fontWeight) {
		configurer.fontWeight(fontWeight);
		return this;
	}

	@Override
	public ExcelCellStyle wrapText(boolean wrapText) {
		configurer.wrapText(wrapText);
		return this;
	};

	@Override
	public void apply(CellStyle cellStyle, Font font) {
		configurer.configure(cellStyle, font);
	}
}
