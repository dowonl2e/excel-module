package com.study.excel.configurer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import com.study.excel.style.align.ExcelAlign;
import com.study.excel.style.align.NoExcelAlign;
import com.study.excel.style.background.ExcelBackgroundColor;
import com.study.excel.style.background.ExcelBackgroundColorStyle;
import com.study.excel.style.background.NoExcelColor;
import com.study.excel.style.border.ExcelBorder;
import com.study.excel.style.border.NoExcelBorder;
import com.study.excel.style.font.ExcelFont;
import com.study.excel.style.font.ExcelFontColorStyle;
import com.study.excel.style.font.NoExcelFontColor;
import com.study.excel.style.font.NoExcelFontWeight;

public final class ExcelCellStyleConfigurer {

	private ExcelAlign align = new NoExcelAlign();
	private ExcelBorder border = new NoExcelBorder();
	private ExcelBackgroundColor foregroundColor = new NoExcelColor();
	private ExcelFont fontColor = new NoExcelFontColor();
	private ExcelFont fontWeight = new NoExcelFontWeight();
	private boolean wrapText = false;
	
	public ExcelCellStyleConfigurer() {}

	public void align(ExcelAlign excelAlign) {
		this.align = excelAlign;
	}

	public void border(ExcelBorder border) {
		this.border = border;
	}
	
	public void foregroundColor(int red, int blue, int green) {
		this.foregroundColor = ExcelBackgroundColorStyle.rgb(red, blue, green);
	}
	
	public void fontColor(int red, int blue, int green) {
		this.fontColor = ExcelFontColorStyle.rgb(red, blue, green);
	}

	public void fontWeight(ExcelFont fontWeight) {
		this.fontWeight = fontWeight;
	}
	
	public void wrapText(boolean wrapText) {
		this.wrapText = wrapText;
	}

	public void configure(CellStyle cellStyle, Font font) {
		align.apply(cellStyle);
		border.apply(cellStyle);
		foregroundColor.applyForegroundWithRGB(cellStyle);
		fontColor.apply(font);
		fontWeight.apply(font);
		cellStyle.setFont(font);
		cellStyle.setWrapText(wrapText);
	}
}
