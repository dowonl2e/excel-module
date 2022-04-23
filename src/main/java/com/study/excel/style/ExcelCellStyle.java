package com.study.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

import com.study.excel.style.align.ExcelAlign;
import com.study.excel.style.border.ExcelBorder;
import com.study.excel.style.font.ExcelFont;

public interface ExcelCellStyle {
	
	ExcelCellStyle align(ExcelAlign align);
	
	ExcelCellStyle border(ExcelBorder border);
	
	ExcelCellStyle foregroundColor(int red, int green, int blue);
	
	ExcelCellStyle fontColor(int red, int green, int blue);
	
	ExcelCellStyle fontWeight(ExcelFont fontWeight);
	
	ExcelCellStyle wrapText(boolean wrapText);
	
	void apply(CellStyle cellStyle, Font font);
	
}
