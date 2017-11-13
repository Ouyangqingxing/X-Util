package com.jason.util.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excel导出工具类
 * @author Jason
 */
public class ExcelExporter {
    
	/**
	 * 将List<Obeject>输出到Excel
	 * @param filePath 输出的绝对路径
	 * @param oList 对象集合
	 * @param illustrate 对象属性描述
	 * @return
	 */
	public boolean getExcel(String filePath , List<Object> oList ,String[] illustrate ){
		FileOutputStream out = null;
		try {
			//1.将对象集合转为字符串集合	2.取得行数与列数
			List<String> strList = getStringList(oList);
			int rowNumber = oList.size();
			int columnNumber = getValueNumber(oList.get(0));
			XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File(filePath)));
			XSSFCellStyle style = workbook.createCellStyle();
			style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			XSSFFont curFont = workbook.createFont();                //设置字体
	        curFont.setFontHeightInPoints((short)90);                //字体大小
	        curFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);         //加粗
	        curFont.setColor((short) 1);
	        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(workbook, 100);
	        Sheet firstSheet = sxssfWorkbook.getSheetAt(0);
	        //3.依次输出表头、表数据
			getSheetTitle(firstSheet , columnNumber , illustrate);
			getSheetData(firstSheet , rowNumber , columnNumber , strList);
			out = new FileOutputStream(filePath);
	        sxssfWorkbook.write(out);
	        return true;
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally{
			if(out != null){
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
	/**
	 * 输出表头
	 * @param firstSheet
	 * @param columnNumber
	 * @param illustrate
	 */
	private void getSheetTitle(Sheet firstSheet , int columnNumber , String[] illustrate){
		Row firstRow =  firstSheet.createRow(0);
		for (int i = 0; i < columnNumber ; i++) {
            CellUtil.createCell( firstRow, i , illustrate[i]);
        }
	}
	
	/**
	 * 输出数据
	 */
	private void getSheetData(Sheet firstSheet , int rowNumber , int columnNumber , List<String> strList){
		for (int i = 1 ; i < rowNumber ; i ++) {
            Row row =  firstSheet.createRow(i);
            for (int j = 0 ; j < columnNumber ; j ++) {
                CellUtil.createCell( row , j ,  strList.get( j + (i-1)*3 ));
            }
        }
	}
	
	/**
	 * 通过反射获取将对象集合转为字符串自合
	 * 输入：List<Object>
	 * 输出：List<Strimg>
	 * @param oList
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	private List<String> getStringList(List<Object> oList) throws IllegalArgumentException, IllegalAccessException{
		List<String> result = new ArrayList<String>();		
		for(int indexOfRow = 0 ; indexOfRow < oList.size() ; indexOfRow++){
			Object object = oList.get(indexOfRow);
			for (Field field : object.getClass().getDeclaredFields()) {
				field.setAccessible(true);
				result.add(field.get(object).toString());
			}
		}
		return result;
	}
	
	/**
	 * 获取对象所拥有的属性数
	 * @return
	 */
	private int getValueNumber(Object object){
		return object.getClass().getDeclaredFields().length;
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}