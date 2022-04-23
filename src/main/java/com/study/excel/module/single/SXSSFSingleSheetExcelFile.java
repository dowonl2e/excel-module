package com.study.excel.module.single;

import java.util.List;

import org.apache.poi.ss.util.CellRangeAddress;

import com.study.excel.data.ExcelData;
import com.study.excel.module.ExcelFile;
import com.study.excel.module.SXSSFExcelFile;

public class SXSSFSingleSheetExcelFile<T> extends SXSSFExcelFile<T> {

	public SXSSFSingleSheetExcelFile() {
		super();
	}

	/**
	 * 시트 생성
	 * 
	 * @param String sheetName
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> createSheet(String sheetName) {
		createSXXSFSheet(sheetName);
		return this;
	}

	/**
	 * 시트 1개 생성 고정
	 * 
	 * @param String sheetName
	 * @param int    sheetCount
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> createSheet(String sheetName, int sheetCount) {
		return createSheet(sheetName);
	}

	/**
	 * 셀을 합병 인덱스 저장 (시작행, 마지막행, 시작열, 마지막열)
	 * 
	 * @param int startRowIndex
	 * @param int endRowIndex
	 * @param int startColumnIndex
	 * @param int endColumnIndex
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> addMergedRegion(int startRowIndex, int endRowIndex, int startColIndex, int endColIndex) {
		startRowIndex = startRowIndex < 0 ? 0 : startRowIndex;
		endRowIndex = endRowIndex < 0 ? 0 : endRowIndex;
		startColIndex = startColIndex < 0 ? 0 : startColIndex;
		endColIndex = endColIndex < 0 ? 0 : endColIndex;

		sheet.addMergedRegion(new CellRangeAddress(startRowIndex, endRowIndex, startColIndex, endColIndex));
		addMergedRegionIndexes(startRowIndex, endRowIndex, startColIndex, endColIndex);
		return this;
	}

	/**
	 * 헤더 데이터 추가
	 * 
	 * @param List<T> headers
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> renderRowHeader(List<? extends Object> headers) {
		renderSXSSFRowHeader(headers);
		return this;
	}

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
	 * 시트 빈 행 추가
	 * 
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> renderEmptyRow() {
		increaseRowIndex();
		return this;
	}

	/**
	 * 행에 셀 데이터 출력
	 * 
	 * @param List<T> data
	 */
	@Override
	public void renderRowData(List<? extends Object> data) {
		renderSXSSFRowData(data);
	}

	@Override
	public void renderRowData(T data) {
		if (data instanceof ExcelData) {
			ExcelData<?> datavo = (ExcelData<?>) data;
			if (datavo != null)
				renderRowData(datavo.getDataList());
		}
	}

	/**
	 * Do Nothing
	 */
	@Override
	public ExcelFile<T> addRowHeader(List<? extends Object> headers) {
		/* Do Nothing */
		return this;
	}

	/**
	 * Do Nothing
	 */
	@Override
	public ExcelFile<T> addRowHeader(T headers) {
		/* Do Nothing */
		return this;
	}

	/**
	 * Do Nothing
	 */
	@Override
	public void renderRowHeader() {
		/* Do Nothing */
	}

}
