package com.jason.util.excel;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * 数据格式转换器
 * @author Jason
 */
public class DataFormatConverter {

	/**
	 * 将 对象集合 转换为 字符串集合
	 * 注意：1.当Object的某属性为空时转为短横线'-'
	 * 	   2.前两个String分别保存rowNumber、columnNumber
	 * @date 2017年11月10日 上午11:15:14
	 * @author Jason
	 * @param objList
	 * @return
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	public static List<String> formateData(List<Object> objList) throws IllegalArgumentException, IllegalAccessException{
		List<String> result = new ArrayList<String>();		
		int rowNumber = objList.size();
		int columnNumber = objList.get(0).getClass().getDeclaredFields().length;
		result.add(rowNumber + "");
		result.add(columnNumber + "");
		for(int indexOfRow = 0 ; indexOfRow < objList.size() ; indexOfRow++){
			Object object = objList.get(indexOfRow);
			for (Field field : object.getClass().getDeclaredFields()) {
				field.setAccessible(true);
				if( null != field && null != field.get(object)){
					result.add(field.get(object).toString());
				}
				else{
					result.add("-");
				}
			}
		}
		return result;
	}
	
}