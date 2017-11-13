package com.jason.test;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Font;

import com.jason.util.excel.DataFormatConverter;
import com.jason.util.excel.ExcelDownloader;
import com.jason.util.excel.ExcelFactory;
import com.jason.util.excel.HeadParam;
import com.jason.util.excel.StyleParam;

/**
 * Excel测试类
 * @author Jason
 */
public class ExcelUtilTest {
	
	/**
	 * main方法  相当于实际项目中的serviceImpl
	 * @date 2017年11月13日 下午3:25:35
	 * @author Jaosn
	 * @param args
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	public static void main(String[] args) throws IllegalArgumentException, IllegalAccessException {
		//1.将对象集合转为字符串集合
		List<Object> objList = getData();
		List<String> strList = DataFormatConverter.formateData(objList);
		//2.设置所需样式，获取Excel(workbook)
		StyleParam sParam = getStyle();
		HSSFWorkbook workbook = ExcelFactory.getHSSFWorkbook(strList,sParam);
		//3.下载Excel
		String filePath = "C:\\Users\\Administrator\\Desktop\\test.xls";
		ExcelDownloader.localDownload(filePath , workbook);
	}
	
	/**
	 * 获取数据
	 * @date 2017年11月13日 下午3:26:54
	 * @author Jason
	 * @return
	 */
	public static List<Object> getData(){
		Student s1 = new Student(1,"学生A",11);
		Student s2 = new Student(2,"学生B",22);
		Student s3 = new Student(3,"学生C",33);
		Student s4 = new Student(4,"学生D",44);
		Student s5 = new Student(5,"学生E",55);
		Student s6 = new Student(6,"学生F",66);
		Student s7 = new Student(7,"学生G",77);
		Student s8 = new Student(8,"学生H",88);
		Student s9 = new Student(9,"学生I",99);
		List<Object> studentList = new ArrayList<Object>();
		studentList.add(s1);
		studentList.add(s2);
		studentList.add(s3);
		studentList.add(s4);
		studentList.add(s5);
		studentList.add(s5);
		studentList.add(s6);
		studentList.add(s7);
		studentList.add(s8);
		studentList.add(s9);
		return studentList;
	}
	
	/**
	 * 获取样式 每次使用该工具只需修改此方法
	 * @date 2017年11月13日 下午3:27:33
	 * @author Jason
	 * @return
	 */
	public static StyleParam getStyle(){
		StyleParam sParam = new StyleParam();
		//设置表头内容及格式
		List<HeadParam> headParamList = new ArrayList<HeadParam>();
		HeadParam num = new HeadParam("编号",100);
		HeadParam name = new HeadParam("姓名",300);
		HeadParam age = new HeadParam("年龄",300);
		headParamList.add(num);
		headParamList.add(name);
		headParamList.add(age);
		//设置表头样式
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFCellStyle headCellStyle = workbook.createCellStyle();
		HSSFFont headerFont = workbook.createFont();
		headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		headerFont.setFontHeightInPoints((short) 10);
		headCellStyle.setFont(headerFont);
		headCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		headCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		headCellStyle.setWrapText(true);
		//设置表格样式
		HSSFCellStyle bodyCellStyle = workbook.createCellStyle();
		bodyCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		bodyCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		bodyCellStyle.setWrapText(true);
		
		sParam.setWorkbook(workbook);
		sParam.setHeadParamList(headParamList);
		sParam.setHeadCellStyle(headCellStyle);
		sParam.setBodyCellStyle(bodyCellStyle);
		return sParam;
	}
	
}