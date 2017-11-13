package com.jason.test;

import java.util.ArrayList;
import java.util.List;

import com.jason.util.excel.ExcelExporter;

public class ExcelUtilTest {
	
	public static void main(String[] args) throws IllegalArgumentException, IllegalAccessException {
		String filePath = "C:\\Users\\Administrator\\Desktop\\test.xlsx";
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
		
		String[] title = {"id","姓名","年龄"};
		
		ExcelExporter ee = new ExcelExporter();
		ee.getExcel(filePath, studentList , title );
		
	}
	
}