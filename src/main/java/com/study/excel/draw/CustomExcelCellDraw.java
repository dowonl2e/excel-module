package com.study.excel.draw;

public class CustomExcelCellDraw {

}
//public class CustomExcelCellDraw implements ExcelCellDraw {
//
//	private ExcelCellCommentConfigurer commentConfigurer = new ExcelCellCommentConfigurer();
//	private ExcelCellPictureConfigurer pictureConfigurer = new ExcelCellPictureConfigurer();
//
//	public CustomExcelCellDraw comment(int startColIndex, int startRowIndex, int endColIndex, int endRowIndex, String commentStr) {
//		commentConfigurer.comment(startColIndex, startRowIndex, endColIndex, endRowIndex, commentStr);
//		return this;
//	}
//
//	public CustomExcelCellDraw picture(int startColIndex, int startRowIndex, int endColIndex, int endRowIndex, String commentStr) {
//		commentConfigurer.comment(startColIndex, startRowIndex, endColIndex, endRowIndex, commentStr);
//		return this;
//	}
//
//	@Override
//	public void applyComment(Drawing drawing) {
//		commentConfigurer.configure(drawing);
//	}
//
//	@Override
//	public void applyComment(Drawing drawing, Font font) {
//		commentConfigurer.configure(drawing, font);
//	}
//
//	@Override
//	public void applyPicture(Drawing drawing) {
//
//	}
//
//}
