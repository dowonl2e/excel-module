package com.study.excel.module;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.SpreadsheetVersion;

import com.study.excel.style.ExcelCellStyle;

import lombok.NonNull;

public interface ExcelFile<T> {

	SpreadsheetVersion excelVersion = SpreadsheetVersion.EXCEL2007;

	/**
	 * 시트 생성
	 * 
	 * @param String sheetName
	 * @return ExcelFile
	 */
	ExcelFile<T> createSheet(String sheetName);

	/**
	 * 시트 N개 생성
	 * 
	 * @param String sheetName
	 * @param int    sheetCount
	 * @return ExcelFile
	 */
	ExcelFile<T> createSheet(String sheetName, int sheetCount);

	/**
	 * 커멘트 그리기
	 * 
	 * @param @NonNull ExcelCellDraw draw
	 */
	// void setCommentDraw(@NonNull ExcelCellDraw draw);

	/**
	 * 셀 합병 영역 설정
	 * 
	 * @param int startRowIndex
	 * @param int endRowIndex
	 * @param int startColumnIndex
	 * @param int endColumnIndex
	 * @return ExcelFile
	 */
	ExcelFile<T> addMergedRegion(int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex);

	/**
	 * 셀 너비 설정
	 * 
	 * @param int columnIndex
	 * @param int width
	 * @return ExcelFile
	 */
	ExcelFile<T> setColumnWidth(int columnIndex, int width);

	/**
	 * 행 높이 설정
	 * 
	 * @param int   rowIndex
	 * @param float height
	 * @return ExcelFile
	 */
	ExcelFile<T> setRowHeight(int rowIndex, float height);

	/**
	 * 헤더 공통 스타일 설정
	 * 
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	ExcelFile<T> setHeaderCellStyle(@NonNull ExcelCellStyle style);

	/**
	 * 지정한 열의 헤더 스타일 적용
	 * 
	 * @param int      columnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	ExcelFile<T> setHeaderCellStyle(int columnIndex, @NonNull ExcelCellStyle style);

	/**
	 * 지정한 행/열의 헤더 스타일 적용
	 * 
	 * @param int      rowIndex
	 * @param int      columnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	ExcelFile<T> setHeaderCellStyle(int rowIndex, int columnIndex, @NonNull ExcelCellStyle style);

	/**
	 * 지정한 열 범위 헤더 스타일 적용
	 * 
	 * @param int      startColumnIndex
	 * @param int      endColumnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	ExcelFile<T> setHeaderRangeCellStyle(int startColumnIndex, int endColumnIndex, @NonNull ExcelCellStyle style);

	/**
	 * 지정한 행의 열 범위 헤더 스타일 적용
	 * 
	 * @param int      rowIndex
	 * @param int      startColumnIndex
	 * @param int      endColumnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	ExcelFile<T> setHeaderRangeCellStyle(int rowIndex, int startColumnIndex, int endColumnIndex, @NonNull ExcelCellStyle style);

	/**
	 * 데이터 공통 스타일 설정
	 * 
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	ExcelFile<T> setDataCellStyle(@NonNull ExcelCellStyle style);

	/**
	 * 지정한 열의 데이터 스타일 적용
	 * 
	 * @param int      columnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	ExcelFile<T> setDataCellStyle(int columnIndex, @NonNull ExcelCellStyle style);

	/**
	 * 헤더 출력
	 * 
	 * @param List<T> headers
	 * @return ExcelFile
	 */
	ExcelFile<T> addRowHeader(List<? extends Object> headers);

	/**
	 * 헤더 출력
	 * 
	 * @param List<T> headers
	 * @return ExcelFile
	 */
	ExcelFile<T> addRowHeader(T headers);

	/**
	 * 헤더 출력
	 */
	void renderRowHeader();

	/**
	 * 헤더 출력
	 * 
	 * @param List<T> headers
	 * @return ExcelFile
	 */
	ExcelFile<T> renderRowHeader(List<? extends Object> headers);

	/**
	 * 헤더 출력
	 * 
	 * @param List<T> headers
	 * @return ExcelFile
	 */
	ExcelFile<T> renderRowHeader(T headers);

	/**
	 * 행에 셀 데이터 출력
	 * 
	 * @param List<T> data
	 */
	void renderRowData(List<? extends Object> data);

	/**
	 * 행에 셀 데이터 출력
	 * 
	 * @param List<T> data
	 */
	void renderRowData(T data);

	/**
	 * 데이터 여부 체크
	 * 
	 * @return 데이터가 없으면 true, 있으면 false
	 */
	boolean isEmptyData();

	/**
	 * 시트에 행 추가여부 확인
	 * 
	 * @return 추가 불가능하면 true, 추가 가능하면 false
	 */
	boolean isFullSheet();

	/**
	 * 시트 빈 행 추가
	 * 
	 * @return ExcelFile
	 */
	ExcelFile<T> renderEmptyRow();

	/**
	 * 최대 생성 시트 개수
	 * 
	 * @param totalCount
	 * @return int
	 */
	int getMaxSheetCount(int headerRowCount, int totalDataRowCount);

	/**
	 * 엑셀 파일 출력
	 * 
	 * @param OutputStream stream
	 * @throws IOException
	 */
	void write(OutputStream stream) throws IOException;

}