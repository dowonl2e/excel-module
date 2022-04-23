package com.study.excel.style.font;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class ExcelFontColorStyle implements ExcelFont {
	
	private static final int MIN_RGB = 0;
	private static final int MAX_RGB = 255;

	private byte red = (byte)MIN_RGB;
	private byte green = (byte)MIN_RGB;
	private byte blue = (byte)MIN_RGB;
	
	private ExcelFontColorStyle(byte red, byte green, byte blue) {
		this.red = red;
		this.green = green;
		this.blue = blue;
	}

	public static ExcelFontColorStyle rgb(int red, int green, int blue) {
		
		red = red < MIN_RGB ? MIN_RGB : red;
		green = green < MIN_RGB ? MIN_RGB : green; 
		blue = blue < MIN_RGB ? MIN_RGB : blue;

		red = red > MAX_RGB ? MAX_RGB : red;
		green = green > MAX_RGB ? MAX_RGB : green;
		blue = blue > MAX_RGB ? MAX_RGB : blue;
		
		return new ExcelFontColorStyle((byte) red, (byte) green, (byte) blue);
	}
	
	@Override
	public void apply(Font font) {
		XSSFFont xssfFont = (XSSFFont)font;
		xssfFont.setColor(new XSSFColor(new byte[]{red, green, blue}));
	}

}
