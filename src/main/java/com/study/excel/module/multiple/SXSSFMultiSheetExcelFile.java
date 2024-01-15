package com.study.excel.module.multiple;

import com.study.excel.data.ExcelData;
import com.study.excel.module.ExcelFile;
import com.study.excel.module.SXSSFExcelFile;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.ArrayList;
import java.util.List;

public class SXSSFMultiSheetExcelFile<T> extends SXSSFExcelFile<T> {

	private List<List<? extends Object>> headersList = new ArrayList<List<? extends Object>>();
	private int dataStartRowIndex = 0;

	public SXSSFMultiSheetExcelFile() {
		super();
	}

	/**
	 * 시트 생성
	 * 
	 * @param sheetName
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> createSheet(String sheetName) {
		return createSheet(sheetName, 1);
	}

	/**
	 * 시트 N개 생성
	 * 
	 * @param sheetName
	 * @param    sheetCount
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> createSheet(String sheetName, int sheetCount) {
		int sheetIndex = 0;
		while (sheetIndex < sheetCount) {
			createSXXSFSheet(sheetName + (sheetIndex + 1));
			sheetIndex++;
		}
		return this;
	}

	/**
	 * 셀을 합병 인덱스 저장 (시작행, 마지막행, 시작열, 마지막열)
	 * 
	 * @param startRowIndex
	 * @param endRowIndex
	 * @param startColIndex
	 * @param endColIndex
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> addMergedRegion(int startRowIndex, int endRowIndex, int startColIndex, int endColIndex) {
		startRowIndex = startRowIndex < 0 ? 0 : startRowIndex;
		endRowIndex = endRowIndex < 0 ? 0 : endRowIndex;
		startColIndex = startColIndex < 0 ? 0 : startColIndex;
		endColIndex = endColIndex < 0 ? 0 : endColIndex;

		int sheetIndex = 0;
		while (sheetIndex < workbook.getNumberOfSheets()) {
			SXSSFSheet multiSheet = (SXSSFSheet) workbook.getSheetAt(sheetIndex);
			multiSheet.addMergedRegion(new CellRangeAddress(startRowIndex, endRowIndex, startColIndex, endColIndex));
			sheetIndex++;
		}
		addMergedRegionIndexes(startRowIndex, endRowIndex, startColIndex, endColIndex);
		return this;
	}

	/**
	 * 행 헤더 데이터 저장
	 * 
	 * @param headers
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> addRowHeader(List<? extends Object> headers) {
		headersList.add(headers);
		return this;
	}

	/**
	 * 행 헤더 데이터 저장
	 * 
	 * @param headers
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> addRowHeader(T headers) {
		if (headers instanceof ExcelData) {
			ExcelData<?> datavo = (ExcelData<?>) headers;
			if (datavo != null)
				addRowHeader(datavo.getDataList());
		}
		return this;
	}

	/**
	 * 행 헤더 데이터 출력
	 * 
	 * @param headers
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> renderRowHeader(List<? extends Object> headers) {
		this.headersList.add(headers);
		renderRowHeader();
		return this;
	}

	/**
	 * 행 헤더 데이터 출력
	 * 
	 * @param headers
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> renderRowHeader(T headers) {
		if (headers instanceof ExcelData) {
			ExcelData<?> datavo = (ExcelData<?>) headers;
			if (datavo != null)
				renderRowHeader(datavo.getDataList());
		}
		return this;
	}

	/**
	 * 행 헤더 데이터 출력
	 * 
	 * @return ExcelFile
	 */
	@Override
	public void renderRowHeader() {
		int sheetIndex = 0;
		int sheetCount = getSheetCount();
		while (sheetIndex < sheetCount) {
			setSXSSFSheetAt(sheetIndex);
			initRowIndex();
			int headersIndex = 0;
			while (headersIndex < headersList.size()) {
				renderSXSSFRowHeader(headersList.get(headersIndex));
				headersIndex++;
			}
			sheetIndex++;
		}
		this.dataStartRowIndex = headersList.size();
		setSXSSFSheetAt(0);
	}

	/**
	 * 행 데이터 출력
	 * 
	 * @param data
	 * @return ExcelFile
	 */
	@Override
	public void renderRowData(List<? extends Object> data) {
		setRowIndex(getRowIndex() < dataStartRowIndex ? dataStartRowIndex : getRowIndex());
		if (isFullSheet() == true) {
			setSXSSFSheetAt(workbook.getSheetIndex(sheet) + 1);
			setRowIndex(dataStartRowIndex);
			initFlushNum();
		}
		renderSXSSFRowData(data);
	}

	/**
	 * 행 데이터 출력
	 * 
	 * @param data
	 * @return ExcelFile
	 */
	@Override
	public void renderRowData(T data) {
		if (data instanceof ExcelData) {
			ExcelData<?> datavo = (ExcelData<?>) data;
			if (datavo != null)
				renderRowData(datavo.getDataList());
		}
	}

	/**
	 * 시트 빈 행 추가
	 * 
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> renderEmptyRow() {
		renderRowHeader(new ArrayList<>());
		return this;
	}

}
