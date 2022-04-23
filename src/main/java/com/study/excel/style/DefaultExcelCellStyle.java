package com.study.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import com.study.excel.style.align.ExcelAlign;
import com.study.excel.style.align.ExcelAlignStyle;
import com.study.excel.style.background.ExcelBackgroundColor;
import com.study.excel.style.background.ExcelBackgroundColorStyle;
import com.study.excel.style.border.ExcelBorder;
import com.study.excel.style.border.ExcelBorderStyle;
import com.study.excel.style.border.ExcelBorderValues;
import com.study.excel.style.font.ExcelFont;
import com.study.excel.style.font.ExcelFontColorStyle;
import com.study.excel.style.font.ExcelFontWeightStyle;
import com.study.excel.style.font.ExcelFontWeightValues;

public enum DefaultExcelCellStyle implements ExcelCellStyle {
	
	GREY(ExcelAlignStyle.CENTER_CENTER
			, ExcelBorderStyle.newInstance(ExcelBorderValues.THIN)
			, ExcelBackgroundColorStyle.rgb(212, 212, 212)
			, ExcelFontColorStyle.rgb(0, 0, 0)
			, ExcelFontWeightStyle.newInstance(ExcelFontWeightValues.BOLD)
			, true),

	BASIC(ExcelAlignStyle.LEFT_CENTER
			, ExcelBorderStyle.newInstance(ExcelBorderValues.NONE)
			, ExcelBackgroundColorStyle.rgb(255, 255, 255)
			, ExcelFontColorStyle.rgb(0, 0, 0)
			, ExcelFontWeightStyle.newInstance(ExcelFontWeightValues.NORMAL)
			, false);
	
	private final ExcelAlign align;
	private final ExcelBorder border;
	private final ExcelBackgroundColor foregroundColor;
	private final ExcelFont fontColor;
	private final ExcelFont fontWeight;
	private final boolean wrapText;
	
	DefaultExcelCellStyle(ExcelAlign align, ExcelBorder border, ExcelBackgroundColor foregroundColor, ExcelFont fontColor, ExcelFont fontWeight, boolean wrapText) {
		this.align = align;
		this.border = border;
		this.foregroundColor = foregroundColor;
		this.fontColor = fontColor;
		this.fontWeight = fontWeight;
		this.wrapText = wrapText;
	}

	@Override
	public void apply(CellStyle cellStyle, Font font) {
		align.apply(cellStyle);		
		border.apply(cellStyle);
		foregroundColor.applyForegroundWithRGB(cellStyle);
		fontColor.apply(font);
		fontWeight.apply(font);
		cellStyle.setFont(font);
		cellStyle.setWrapText(wrapText);
	}
	
	@Override
	public DefaultExcelCellStyle align(ExcelAlign align) {
		/* Do Nothing */
		return this;
	}

	@Override
	public DefaultExcelCellStyle border(ExcelBorder border) {
		/* Do Nothing */
		return this;
	}

	@Override
	public DefaultExcelCellStyle foregroundColor(int red, int green, int blue) {
		/* Do Nothing */
		return this;
	}

	@Override
	public DefaultExcelCellStyle fontColor(int red, int green, int blue) {
		/* Do Nothing */
		return this;
	}

	@Override
	public DefaultExcelCellStyle fontWeight(ExcelFont fontWeight) {
		/* Do Nothing */
		return this;
	}

	@Override
	public ExcelCellStyle wrapText(boolean isWrapText) {
		/* Do Nothing */
		return null;
	}
	
}
