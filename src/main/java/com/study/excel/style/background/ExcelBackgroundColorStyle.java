package com.study.excel.style.background;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

public class ExcelBackgroundColorStyle implements ExcelBackgroundColor {
	

	private static final int MIN_RGB = 0;
	private static final int MAX_RGB = 255;

	private byte red = (byte)MAX_RGB;
	private byte green = (byte)MAX_RGB;
	private byte blue = (byte)MAX_RGB;
	private short fillPattern = ExcelFillPatternValues.SOLID.getPattern();
	
	private ExcelBackgroundColorStyle(byte red, byte green, byte blue) {
		validateRGB(red, green, blue);
		this.red = red;
		this.green = green;
		this.blue = blue;
	}

	private ExcelBackgroundColorStyle(byte red, byte green, byte blue, short fillPattern) {
		this(red, green, blue);
		this.fillPattern = fillPattern;
	}
	
	public static ExcelBackgroundColorStyle rgb(int red, int green, int blue) {
		return new ExcelBackgroundColorStyle((byte) red, (byte) green, (byte) blue);
	}
	
	public static ExcelBackgroundColorStyle rgb(int red, int green, int blue, short fillPattern) {
		return new ExcelBackgroundColorStyle((byte) red, (byte) green, (byte) blue, fillPattern);
	}
	
	private void validateRGB(int red, int green, int blue) {
		red = red < MIN_RGB ? MIN_RGB : red;
		green = green < MIN_RGB ? MIN_RGB : green; 
		blue = blue < MIN_RGB ? MIN_RGB : blue;

		red = red > MAX_RGB ? MAX_RGB : red;
		green = green > MAX_RGB ? MAX_RGB : green;
		blue = blue > MAX_RGB ? MAX_RGB : blue;
	}
	
	@Override
	public void applyForegroundWithRGB(CellStyle cellStyle) {
		XSSFCellStyle xssfCellStyle = (XSSFCellStyle)cellStyle;
		xssfCellStyle.setFillForegroundColor(new XSSFColor(new byte[]{red, green, blue}));
		cellStyle.setFillPattern(fillPattern);
	}
	
	@Override
	public void applyBackgroundWithRGB(CellStyle cellStyle) {
		XSSFCellStyle xssfCellStyle = (XSSFCellStyle)cellStyle;
		xssfCellStyle.setFillBackgroundColor(new XSSFColor(new byte[]{red, green, blue}));
		cellStyle.setFillPattern(fillPattern);
	}
}
