package com.jason.util.excel;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 * 样式参数
 * @author Jason
 */
public class StyleParam {

	/**
	 * 表头样式
	 */
	private List<HeadParam> headParamList;
	
	/**
	 * workbook
	 */
	private HSSFWorkbook workbook;
	
	/**
	 * 表头样式
	 */
	private HSSFCellStyle headCellStyle;

	/**
	 * 表格样式
	 */
	private HSSFCellStyle bodyCellStyle;

	public List<HeadParam> getHeadParamList() {
		return headParamList;
	}

	public void setHeadParamList(List<HeadParam> headParamList) {
		this.headParamList = headParamList;
	}

	public HSSFWorkbook getWorkbook() {
		return workbook;
	}

	public void setWorkbook(HSSFWorkbook workbook) {
		this.workbook = workbook;
	}

	public HSSFCellStyle getHeadCellStyle() {
		return headCellStyle;
	}
	

	public void setHeadCellStyle(HSSFCellStyle headCellStyle) {
		this.headCellStyle = headCellStyle;
	}
	

	public HSSFCellStyle getBodyCellStyle() {
		return bodyCellStyle;
	}
	

	public void setBodyCellStyle(HSSFCellStyle bodyCellStyle) {
		this.bodyCellStyle = bodyCellStyle;
	}

}