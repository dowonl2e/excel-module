package com.study.excel.module;

import com.study.excel.style.ExcelCellStyle;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;

import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

@Slf4j
public abstract class SXSSFExcelFile<T> implements ExcelFile<T> {

	/* 기본 시트 */
	protected final static String SIMPLE_SHEETNAME = "Sheet";

	/* 메모리 Flush */
	private static final int FLUSH_COUNT = 1000;

	/* 기본 행 높이 */
	private static final float DEFAULT_ROW_HEIGHT = 17;

	/* 엑셀 */
	protected SXSSFWorkbook workbook;
	protected SXSSFSheet sheet;

	/* 스타일 */
	private XSSFCellStyle headerStyle, dataStyle;
	private XSSFFont headerFont, dataFont;
	private Map<Integer, Object> targetHeaderStyleMap = new HashMap<>();
	private Map<Integer, Object> targetDataStyleMap = new HashMap<>();

	private Map<Integer, Object> rowStyleMap = new HashMap<>();

	/* 행/열/메모리Flush/시트 최대 행 인덱스 초기화 */
	private int rowIndex = 0, colIndex = 0, flushNum = 1, maxColIndex = 0;

	/* 셀 합병 인덱스 */
	private List<Integer> startRowIndexes = new ArrayList<>();
	private List<Integer> endRowIndexes = new ArrayList<>();
	private List<Integer> startColIndexes = new ArrayList<>();
	private List<Integer> endColIndexes = new ArrayList<>();

	private Map<Integer, Float> rowHeightMap = new HashMap<>();
	private Map<Integer, Integer> columnWidthMap = new HashMap<>();

	private List<List<T>> cellDataList = new ArrayList<>();

	protected SXSSFExcelFile() {
		this.workbook = new SXSSFWorkbook();
		initStyles();
	}

	private void initStyles(){
		this.headerStyle = (XSSFCellStyle) workbook.createCellStyle();
		this.headerFont = (XSSFFont) workbook.createFont();
		this.dataStyle = (XSSFCellStyle) workbook.createCellStyle();
		this.dataFont = (XSSFFont) workbook.createFont();
	}

	/**
	 * 행에 데이터 출력
	 * 
	 * @param data
	 */
	protected void renderSXSSFRowData(List<? extends Object> data) {
		if (isFullSheet())
			throw new RuntimeException("ExcelFileRowRangeException : range(0..." + excelVersion.getLastRowIndex() + ") / access(" + getRowIndex() + ")");

		if (data != null && data.size() > 0) {
			SXSSFRow row = (SXSSFRow) sheet.createRow(rowIndex);
			colIndex = 0;
			SXSSFCell cell = null;
			for (Object obj : data) {
				cell = ObjectUtils.isEmpty(row.getCell(colIndex)) ? (SXSSFCell) row.createCell(colIndex) : (SXSSFCell) row.getCell(colIndex);
				cell.setCellStyle(getDataCellStyle(rowIndex, colIndex));
				renderCellValue(cell, (obj == null ? "" : obj));
				colIndex++;
			}
			rowIndex++;
			flush();
		}
	}

	/**
	 * 엑셀 시트 생성 (시트명 값이 Null 또는 빈값("")의 경우 기본 시트명(Sheet)으로 설정됨)
	 * @param sheetName
	 */
	protected void createSXXSFSheet(String sheetName) {
		sheetName = ObjectUtils.isEmpty(sheetName) ? SIMPLE_SHEETNAME : sheetName;
		if (workbook.getSheetIndex(sheetName) > -1)
			sheetName += (getSheetCount() + 1);
		this.sheet = (SXSSFSheet) workbook.createSheet(sheetName);
	}

	/**
	 * 헤더 데이터 추가 (스타일과 셀 합병 설정 후에 호출)
	 * 
	 * @param headers
	 */
	protected void renderSXSSFRowHeader(List<? extends Object> headers) {
		if (!CollectionUtils.isEmpty(headers)) {
			SXSSFRow row = (SXSSFRow) sheet.createRow(rowIndex);
			row.setHeightInPoints(rowHeightMap.get(rowIndex) == null ? DEFAULT_ROW_HEIGHT : rowHeightMap.get(rowIndex));

			colIndex = 0;
			/**
			 * O O O O
			 * O O O O
			 * O O X X
			 * O O O O
			 */
			Queue<? extends Object> headerQueue = new LinkedList<>(headers);
			int cellStartRowIndex = rowIndex, cellStartColIndex = colIndex, cellEndColIndex = colIndex;
			while (!headerQueue.isEmpty()) {
				boolean isMergeResion = false;

				// 행/열 인덱스가 합병 영역에 포함되는지 체크
				for (int idx = 0; idx < startRowIndexes.size(); idx++) {
					int mergeStartRowIdx = startRowIndexes.get(idx);
					int mergeEndRowIdx = endRowIndexes.get(idx);
					int mergeStartColIdx = startColIndexes.get(idx);
					int mergeEndColIdx = endColIndexes.get(idx);
					/*
						추가하려는 셀 행/열이 설정된 합병영역에 시작 행/열 시작 인덱스 체크
					 */
					if (rowIndex >= mergeStartRowIdx && rowIndex <= mergeEndRowIdx
							&& colIndex >= mergeStartColIdx && colIndex <= mergeStartColIdx) {
						cellStartRowIndex = mergeStartRowIdx;
						cellStartColIndex = mergeStartColIdx;
						cellEndColIndex = mergeEndColIdx;
						isMergeResion = true;
						break;
					}
				}

				/*
				  행/열 인덱스가 합병 영역에 포함되는 경우
				  1. 행/열 인덱스가 합병영역의 시작 위치인 경우 데이터 추가 후 다음열부터 합병 영역의 마지막 열까지 빈 셀 추가
				  2. 행/열 인덱스가 합병영역의 시작 위치가 아닌 경우 현재 행에서 합병영역의 마지막 열까지 빈 셀만 추가
				 */
				if (isMergeResion) {
					if (rowIndex == cellStartRowIndex && colIndex == cellStartColIndex) {
						createCell(row, rowIndex, colIndex, headerQueue.poll());
						createEmptyCell(row, rowIndex, colIndex, cellEndColIndex);
					}
					else {
						createEmptyCell(row, rowIndex, colIndex, cellEndColIndex);
					}
				}
				// 행/열 인덱스가 합병 영역에 포함되지 않는 경우
				else {
					createCell(row, rowIndex, colIndex, headerQueue.poll());
				}
			}

			/*
			 * 현재 행에 셀 추가가 끝난 후 현재 열 인덱스가 다른 행의 열 인덱스보다 적은 경우 최대 열위치만큼 빈 셀 추가
			 * (마지막 열이 합병으로 끝나는 경우에 빈셀을 추가하기 위함)
			 */
			maxColIndex = maxColIndex < colIndex ? colIndex : maxColIndex;

			if (colIndex < maxColIndex)
				createEmptyCell(row, rowIndex, colIndex, (maxColIndex - 1));
		}
		rowIndex++;
	}

	/**
	 * 특정 행/열에 저장된 헤더 스타일 반환
	 * 
	 * @param rowIndex
	 * @param columnIndex
	 * @return XSSFCellStyle
	 */
	@SuppressWarnings("unchecked")
	private XSSFCellStyle getHeaderCellStyle(int rowIndex, int columnIndex) {
		if (rowStyleMap.get(rowIndex) != null) {
			Map<Integer, Object> styleMap = (HashMap<Integer, Object>) rowStyleMap.get(rowIndex);
			if (styleMap.get(columnIndex) != null)
				return (XSSFCellStyle) styleMap.get(columnIndex);
			else
				return headerStyle;
		} else if (targetHeaderStyleMap.get(columnIndex) != null) {
			return (XSSFCellStyle) targetHeaderStyleMap.get(columnIndex);
		} else {
			return headerStyle;
		}
	}

	/**
	 * 특정 행/열에 저장된 데이터 스타일 반환
	 * 
	 * @param rowIndex
	 * @param columnIndex
	 * @return XSSFCellStyle
	 */
	@SuppressWarnings("unchecked")
	private XSSFCellStyle getDataCellStyle(int rowIndex, int columnIndex) {
		if (rowStyleMap.get(rowIndex) != null) {
			Map<Integer, Object> styleMap = (HashMap<Integer, Object>) rowStyleMap.get(rowIndex);
			if (styleMap.get(columnIndex) != null)
				return (XSSFCellStyle) styleMap.get(columnIndex);
			else
				return dataStyle;
		} else if (targetDataStyleMap.get(columnIndex) != null) {
			return (XSSFCellStyle) targetDataStyleMap.get(columnIndex);
		} else {
			return dataStyle;
		}
	}

	/**
	 * 셀 데이터 출력
	 * 
	 * @param cell
	 * @param cellValue
	 */
	private void renderCellValue(SXSSFCell cell, Object cellValue) {
		if (cellValue instanceof Number) {
			Number numberValue = (Number) cellValue;
			cell.setCellValue(numberValue.doubleValue());
			return;
		}
		cell.setCellValue(cellValue == null ? "" : cellValue.toString());
	}

	/**
	 * 셀 생성 및 스타일, 값 설정
	 * 
	 * @param row
	 * @param rowIndex
	 * @param columnIndex
	 * @param cellValue
	 */
	private void createCell(SXSSFRow row, int rowIndex, int columnIndex, Object cellValue) {

		int columnWidth = (columnWidthMap.get(columnIndex) == null ? 0 : (Integer) columnWidthMap.get(columnIndex));
		if (columnWidth > 0)
			sheet.setColumnWidth(columnIndex, columnWidth);

		SXSSFCell cell = ObjectUtils.isEmpty(row.getCell(columnIndex))
				? (SXSSFCell) row.createCell(columnIndex) : (SXSSFCell) row.getCell(columnIndex);

		cell.setCellStyle(getHeaderCellStyle(rowIndex, columnIndex));
		if (cellValue != null) {
			renderCellValue(cell, (cellValue == null ? "" : cellValue));
		}
		this.colIndex++;
	}

	/**
	 * 셀 합병에 필요한 빈 셀 생성
	 * 
	 * @param row
	 * @param rowIndex
	 * @param startColumnIndex
	 * @param endColumnIndex
	 */
	private void createEmptyCell(SXSSFRow row, int rowIndex, int startColumnIndex, int endColumnIndex) {
		int idx = startColumnIndex;
		while (idx <= endColumnIndex) {
			createCell(row, rowIndex, idx++, null);
		}
	}

	/**
	 * 메모리 Flush
	 */
	private void flush() {
		if (rowIndex == FLUSH_COUNT * flushNum) {
			try {
				sheet.flushRows(FLUSH_COUNT);
			} catch (IOException e) {
				e.printStackTrace();
			}
			flushNum++;
		}
	}

	/**
	 * 헤더 공통 스타일 설정
	 * 
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> setHeaderCellStyle(@NonNull ExcelCellStyle cellStyle) {
		cellStyle.apply(headerStyle, headerFont);
		return this;
	}

	/**
	 * 지정한 열의 헤더 스타일 적용
	 * 
	 * @param columnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> setHeaderCellStyle(int columnIndex, @NonNull ExcelCellStyle cellStyle) {
		if (columnIndex >= 0) {
			XSSFCellStyle tempStyle = (XSSFCellStyle) workbook.createCellStyle();
			cellStyle.apply(tempStyle, (XSSFFont) workbook.createFont());
			if (targetHeaderStyleMap.get(columnIndex) != null)
				targetHeaderStyleMap.remove(columnIndex);
			targetHeaderStyleMap.put(columnIndex, tempStyle);
		}
		return this;
	}

	/**
	 * 지정한 행/열의 헤더 스타일 적용
	 * 
	 * @param rowIndex
	 * @param columnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	@SuppressWarnings("unchecked")
	@Override
	public ExcelFile<T> setHeaderCellStyle(int rowIndex, int columnIndex, @NonNull ExcelCellStyle cellStyle) {
		if (rowIndex >= 0 && columnIndex >= 0) {
			XSSFCellStyle tempStyle = (XSSFCellStyle) workbook.createCellStyle();
			cellStyle.apply(tempStyle, (XSSFFont) workbook.createFont());

			Map<Integer, Object> colHeaderStyleMap = null;
			if (rowStyleMap.get(rowIndex) == null)
				colHeaderStyleMap = new HashMap<Integer, Object>();
			else
				colHeaderStyleMap = (HashMap<Integer, Object>) rowStyleMap.get(rowIndex);

			if (colHeaderStyleMap.get(columnIndex) != null)
				colHeaderStyleMap.remove(columnIndex);
			colHeaderStyleMap.put(columnIndex, tempStyle);

			if (rowStyleMap.get(rowIndex) != null)
				rowStyleMap.remove(rowIndex);
			rowStyleMap.put(rowIndex, colHeaderStyleMap);
		}
		return this;
	}

	/**
	 * 지정한 열 범위 헤더 스타일 적용
	 * 
	 * @param startColumnIndex
	 * @param endColumnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> setHeaderRangeCellStyle(int startColumnIndex, int endColumnIndex,
			@NonNull ExcelCellStyle cellStyle) {
		if (startColumnIndex >= 0 && endColumnIndex >= 0) {
			XSSFCellStyle tempStyle = (XSSFCellStyle) workbook.createCellStyle();
			cellStyle.apply(tempStyle, (XSSFFont) workbook.createFont());

			int idx = startColumnIndex;
			while (idx < endColumnIndex) {
				if (targetHeaderStyleMap.get(idx) != null)
					targetHeaderStyleMap.remove(idx);
				targetHeaderStyleMap.put(idx, tempStyle);
				idx++;
			}
		}
		return this;
	}

	/**
	 * 지정한 행의 열 범위 헤더 스타일 적용
	 * 
	 * @param rowIndex
	 * @param startColumnIndex
	 * @param endColumnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	@SuppressWarnings("unchecked")
	@Override
	public ExcelFile<T> setHeaderRangeCellStyle(int rowIndex, int startColumnIndex, int endColumnIndex,
			@NonNull ExcelCellStyle cellStyle) {
		if (rowIndex >= 0 && startColumnIndex >= 0 && endColumnIndex >= 0) {
			XSSFCellStyle tempStyle = (XSSFCellStyle) workbook.createCellStyle();
			cellStyle.apply(tempStyle, (XSSFFont) workbook.createFont());

			Map<Integer, Object> colHeaderStyleMap = null;
			if (rowStyleMap.get(rowIndex) == null)
				colHeaderStyleMap = new HashMap<Integer, Object>();
			else
				colHeaderStyleMap = (HashMap<Integer, Object>) rowStyleMap.get(rowIndex);

			int idx = startColumnIndex;
			while (idx < endColumnIndex) {
				if (colHeaderStyleMap.get(idx) != null)
					colHeaderStyleMap.remove(idx);
				colHeaderStyleMap.put(idx, tempStyle);
				idx++;
			}

			if (rowStyleMap.get(rowIndex) != null)
				rowStyleMap.remove(rowIndex);
			rowStyleMap.put(rowIndex, colHeaderStyleMap);
		}
		return this;
	}

	/**
	 * 데이터 공통 스타일 설정
	 * 
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> setDataCellStyle(@NonNull ExcelCellStyle cellStyle) {
		cellStyle.apply(dataStyle, dataFont);
		return this;
	}

	/**
	 * 지정한 열의 데이터 스타일 적용
	 * 
	 * @param columnIndex
	 * @param @NonNull ExcelCellStyle style
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> setDataCellStyle(int columnIndex, @NonNull ExcelCellStyle cellStyle) {
		if (columnIndex >= 0) {
			XSSFCellStyle tempStyle = (XSSFCellStyle) workbook.createCellStyle();
			cellStyle.apply(tempStyle, (XSSFFont) workbook.createFont());
			targetDataStyleMap.put(columnIndex, tempStyle);
		}
		return this;
	}

	/**
	 * 셀을 합병 인덱스 저장 (시작행, 마지막행, 시작열, 마지막열)
	 * 
	 * @param startRowIndex
	 * @param endRowIndex
	 * @param startColumnIndex
	 * @param endColumnIndex
	 * @return ExcelFile
	 */
	protected void addMergedRegionIndexes(int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex) {
		startRowIndexes.add(startRowIndex);
		endRowIndexes.add(endRowIndex);
		startColIndexes.add(startColumnIndex);
		endColIndexes.add(endColumnIndex);


	}

	/**
	 * 출력 데이터 빈값 여부 확인
	 * 
	 * @return 데이터가 없으면 true, 있으면 false
	 */
	@Override
	public boolean isEmptyData() {
		return (this.cellDataList == null || this.cellDataList.size() == 0) ? true : false;
	}

	/**
	 * 행 인덱스 초기화
	 */
	protected void initRowIndex() {
		rowIndex = 0;
	}

	/**
	 * 행 인덱스 증가(+1)
	 */
	protected void increaseRowIndex() {
		this.rowIndex++;
	}

	/**
	 * 행 인덱스 변경
	 * 
	 * @param rowIndex
	 */
	protected void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	/**
	 * 현재 행 인덱스 반환
	 * 
	 * @return int
	 */
	protected int getRowIndex() {
		return this.rowIndex;
	}

	/**
	 * 시트에 행 추가여부 확인
	 * 추가 불가능하면 true, 추가 가능하면 false
	 * @return boolean
	 */
	public boolean isFullSheet() {
		return this.rowIndex > excelVersion.getLastRowIndex();
	}

	/**
	 * 메모리 flush 값 초기화(시트 변경시 사용)
	 */
	protected void initFlushNum() {
		flushNum = 1;
	}

	/**
	 * 시트 변경
	 * 
	 * @param sheetIndex
	 */
	protected void setSXSSFSheetAt(int sheetIndex) {
		if (sheetIndex > getSheetCount())
			throw new RuntimeException(
					"SheetIndexOutOfRangeException : max(" + getSheetCount() + ") / access(" + sheetIndex + ")");

		this.sheet = (SXSSFSheet) workbook.getSheetAt(sheetIndex);
	}

	/**
	 * 최대 생성 시트 개수
	 * 
	 * @param headerRowCount
	 * @param totalDataRowCount
	 * @return int
	 */
	@Override
	public int getMaxSheetCount(int headerRowCount, int totalDataRowCount) {
		int div = (headerRowCount + totalDataRowCount) / excelVersion.getMaxRows();
		int mod = (headerRowCount + totalDataRowCount) % excelVersion.getMaxRows();
		int sheetCount = (mod == 0 ? div : (div + 1));
		return sheetCount;
	}

	/**
	 * 생성된 시트 개수
	 * 
	 * @return int
	 */
	protected int getSheetCount() {
		return workbook == null ? 0 : workbook.getNumberOfSheets();
	}

	/**
	 * 행 높이 설정
	 * 
	 * @param rowIndex
	 * @param height
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> setRowHeight(int rowIndex, float height) {
		rowHeightMap.put(rowIndex, height);
		return this;
	}

	/**
	 * 열 넓이 설정
	 * 
	 * @param columnIndex
	 * @param width
	 * @return ExcelFile
	 */
	@Override
	public ExcelFile<T> setColumnWidth(int columnIndex, int width) {
		columnWidthMap.put(columnIndex, width);
		return this;
	}

	/**
	 * 엑셀 출력
	 * 
	 * @param stream
	 * @throws IOException
	 */
	@Override
	public void write(OutputStream stream) throws IOException {
		workbook.write(stream);
		workbook.dispose();
		stream.flush();
		if (stream != null)
			stream.close();
	}

}