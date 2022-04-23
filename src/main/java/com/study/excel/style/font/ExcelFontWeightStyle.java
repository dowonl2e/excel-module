package com.study.excel.style.font;

import org.apache.poi.ss.usermodel.Font;

public class ExcelFontWeightStyle implements ExcelFont {

	private ExcelFontWeightValues weight;
	
	private ExcelFontWeightStyle(ExcelFontWeightValues weight) {
		this.weight = weight;
	}
	
	public static ExcelFontWeightStyle newInstance(ExcelFontWeightValues weight) {
		return new ExcelFontWeightStyle(weight == null ? ExcelFontWeightValues.NORMAL : weight);
	}
	
	@Override
	public void apply(Font font) {
		font.setBoldweight(weight.getWeight());
	}

}
