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
	 * 将List<Object>转换为List<String>
	 * 当Object的某属性为空时转为短横线'-'
	 * @date 2017年11月10日 上午11:15:14
	 * @author Jason
	 * @param objList
	 * @return
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	public List<String> formateData(List<Object> objList) throws IllegalArgumentException, IllegalAccessException{
		List<String> result = new ArrayList<String>();		
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