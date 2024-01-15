package com.study.controller;

import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import com.study.excel.data.ExcelData;
import com.study.excel.module.ExcelFile;
import com.study.excel.module.multiple.SXSSFMultiSheetExcelFile;
import com.study.excel.module.single.SXSSFSingleSheetExcelFile;
import com.study.excel.style.CustomExcelCellStyle;
import com.study.excel.style.DefaultExcelCellStyle;
import com.study.excel.style.ExcelCellStyle;
import com.study.excel.style.align.ExcelAlignStyle;
import com.study.excel.style.border.ExcelBorderStyle;
import com.study.excel.style.border.ExcelBorderValues;

@Controller
public class ExcelTestController {

	@GetMapping("/down1")
	public void down1(HttpServletRequest request, HttpServletResponse response) throws IOException {
		
		setFileNameToResponse(request, response, "ExcelModuleTest");
		
		//기본 헤더 스타일
		ExcelCellStyle headerStyle = DefaultExcelCellStyle.GREY;
		
		//커스텀 스타일
		ExcelCellStyle dataStyle = new CustomExcelCellStyle();
		dataStyle.align(ExcelAlignStyle.LEFT_CENTER)
				.border(ExcelBorderStyle.newInstance(ExcelBorderValues.THIN))
				.foregroundColor(255, 255, 255)
				.fontColor(0, 0, 0).wrapText(true);
		
		//헤더 스타일/합병/행,열 크기 지정
		ExcelFile<ExcelData<?>> singleSheetExcel = new SXSSFSingleSheetExcelFile<>();
		singleSheetExcel.createSheet("엑셀 싱글 시트")
						.setHeaderCellStyle(headerStyle)
						.addMergedRegion(0, 2, 0, 0)
						.addMergedRegion(0, 0, 1, 3)
						.addMergedRegion(1, 1, 1, 3)
						.addMergedRegion(3, 3, 0, 3)
						.setRowHeight(0, 30)
						.setColumnWidth(2, 3000)
						.setColumnWidth(3, 4000);
		
		//헤더 추가 
		ExcelData<String> headerData = new ExcelData<>();
		headerData.addData("No.","정보");
		singleSheetExcel.renderRowHeader(headerData);
		headerData.clearData().addData("개인정보");
		singleSheetExcel.renderRowHeader(headerData);
		headerData.clearData().addData("이름","연락처","이메일");
		singleSheetExcel.renderRowHeader(headerData);
		headerData.clearData().addData("아래는 데이터입니다.");
		singleSheetExcel.renderRowHeader(headerData);

		//셀 데이터 추가
		ExcelData<Object> data = null;
		singleSheetExcel.setDataCellStyle(dataStyle);
		for(int i = 0 ; i < 100 ; i++) {
			data = new ExcelData<>();
			data.addData(100-i).addData("이름" +(i+1)).addData("연락처"+(i+1)).addData("이메일"+(i+1));
			singleSheetExcel.renderRowData(data);
		}
		
		singleSheetExcel.write(response.getOutputStream());
		
	}
	
	@GetMapping("/down2")
	public void down2(HttpServletRequest request, HttpServletResponse response) throws IOException {

//		setFileNameToResponse(request, response, "ExcelModuleTest");
//
//		//기본 헤더 스타일
//		ExcelCellStyle headerStyle = DefaultExcelCellStyle.GREY;
//
//		//커스텀 스타일
//		ExcelCellStyle dataStyle = new CustomExcelCellStyle();
//		dataStyle.align(ExcelAlignStyle.LEFT_CENTER)
//				.border(ExcelBorderStyle.newInstance(ExcelBorderValues.THIN))
//				.foregroundColor(255, 255, 255)
//				.fontColor(0, 0, 0);
//
//		int dataCount = 1050000;
//		ExcelFile<Object> multiSheetFile = new SXSSFMultiSheetExcelFile<Object>();
//		int maxSheetCount = multiSheetFile.getMaxSheetCount(2, dataCount);
//
//		multiSheetFile.createSheet("엑셀 멀티 시트", maxSheetCount)
//						.setHeaderCellStyle(headerStyle)
//						.addMergedRegion(0, 1, 0, 0)
//						.addMergedRegion(0, 0, 1, 3)
//						.setRowHeight(0, 30)
//						.setColumnWidth(2, 3000)
//						.setColumnWidth(3, 4000);;
//
//		//헤더 추가
//		ExcelData<String> headerData = new ExcelData<>();
//		headerData.addData("No.").addData("정보");
//		multiSheetFile.addRowHeader(headerData);
//		headerData.clearData().addData("이름").addData("연락처").addData("이메일");
//		multiSheetFile.addRowHeader(headerData).renderRowHeader();
//
//		//데이터 추가
//		ExcelData<Object> data = null;
//		multiSheetFile.setDataCellStyle(dataStyle);
//		for(int i = 0 ; i < dataCount ; i++) {
//			data = new ExcelData<>();
//			data.addData(dataCount-i).addData("이름" +(i+1)).addData("연락처"+(i+1)).addData("이메일"+(i+1));
//			multiSheetFile.renderRowData(data);
//		}
//
//		multiSheetFile.write(response.getOutputStream());
		
	}
	
	private String createFileName(HttpServletRequest request, String fileName) {
		SimpleDateFormat fileFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
		try {
			String userAgent = request.getHeader("User-Agent");
			if (userAgent.contains("Trident") || (userAgent.indexOf("MSIE") > -1)) {
				fileName = URLEncoder.encode(fileName, "UTF-8").replaceAll("\\+", "%20");
			} 
			else if (userAgent.contains("Chrome") || userAgent.contains("Opera") || userAgent.contains("Firefox")) {
				fileName = new String(fileName.getBytes("UTF-8"), "ISO-8859-1");
			}
		}
		catch(UnsupportedEncodingException e) {
			fileName = "Excel";
		}
		return fileName + "_" + fileFormat.format(new Date()) + ".xlsx";
	}
	
	private void setFileNameToResponse(HttpServletRequest request, HttpServletResponse response, String fileName) {
		String encodeFileName = createFileName(request, fileName);
		String userAgent = request.getHeader("User-Agent");
	    if (userAgent.indexOf("MSIE 5.5") >= 0) {
	    	response.setContentType("doesn/matter");
	    	response.setHeader("Content-Disposition", "filename=\"" + encodeFileName + "\"");
	    } else {
	    	response.setHeader("Content-Disposition", "attachment; filename=\"" + encodeFileName + "\"");
	    }
	}
	
}
