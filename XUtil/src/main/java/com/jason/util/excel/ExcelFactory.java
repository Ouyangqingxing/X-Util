package com.jason.util.excel;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * Excel工厂
 * @author Jason
 */
@SuppressWarnings("deprecation")
public class ExcelFactory {
	
	/**
	 * 将 字符串集合 转为 HSSFWorkBook
	 * @date 2017年11月13日 上午9:28:31
	 * @author Jason
	 * @param strList
	 * @return
	 */
	public static HSSFWorkbook getHSSFWorkbook(List<String> strList , StyleParam styleParam){
		HSSFWorkbook workbook = styleParam.getWorkbook();
		//生成表头
		HSSFCellStyle headCellStyle = styleParam.getHeadCellStyle();
		List<HeadParam> headParamList = styleParam.getHeadParamList();
		HSSFSheet sheet = workbook.createSheet();
		generateHead(headCellStyle, sheet, headParamList);
		//生成表格		
		HSSFSheet bodySheet = workbook.getSheetAt(0);
		HSSFCellStyle bodyCellStyle = styleParam.getBodyCellStyle();
		generateBody(bodyCellStyle , bodySheet , strList);
		return workbook;
	}

	/**
	 * 生成表头
	 * @date 2017年11月13日 下午3:17:52
	 * @author Jason
	 * @param headCellStyle
	 * @param sheet
	 * @param headParamList
	 */
	private static void generateHead(HSSFCellStyle headCellStyle,HSSFSheet sheet,List<HeadParam> headParamList) {
		HSSFRow row = sheet.createRow(0);
		int cellNum = 0;
		for (int i = 0 ; i < headParamList.size() ; i++) {
			HeadParam hParam = headParamList.get(i);
			sheet.setColumnWidth(i , (hParam.getWidth() * 20));
			HSSFCell cell = row.createCell(cellNum);
			cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			cell.setCellStyle(headCellStyle);
			sheet.addMergedRegion(new CellRangeAddress(0, hParam.getRowspan() - 1,cellNum, cellNum + hParam.getColumn() - 1));
			cell.setCellValue(hParam.getName());
			cellNum = cellNum + hParam.getColumn();
		}
		cellNum = 1;
	}
	
	/**
	 * 生成表格
	 * @date 2017年11月13日 下午3:17:25
	 * @author Jason
	 * @param bodyCellStyle
	 * @param sheet
	 * @param strList
	 */
	private static void generateBody(HSSFCellStyle bodyCellStyle , HSSFSheet sheet , List<String> strList){
		int rowNumber = Integer.parseInt(strList.get(0));
		int columnNumber = Integer.parseInt(strList.get(1));
		strList.remove(0);
		strList.remove(0);
		for (int i = 1 ; i < rowNumber + 1 ; i ++) {
			HSSFRow row =  sheet.createRow(i);
            for (int j = 0 ; j < columnNumber ; j ++) {
            	HSSFCell cell = row.createCell(j);
    			cell.setCellType(XSSFCell.CELL_TYPE_STRING);
    			cell.setCellStyle(bodyCellStyle);
    			cell.setCellValue(strList.get( j + (i - 1) * columnNumber ));
            }
        }
	}
	
}