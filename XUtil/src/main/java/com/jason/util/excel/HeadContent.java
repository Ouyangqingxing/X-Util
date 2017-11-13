package com.jason.util.excel;

/**
 * 表头参数
 * @author Jason
 */
public class HeadContent {

	/**
	 * 名称
	 */
	private String name;
	
	/**
	 * 宽度
	 */
	private int width;
	
	/**
	 * 跨行数 默认为1
	 */
	private int rowspan;
	
	/**
	 * 跨列数，默认为1
	 */
	private int column;

	public HeadContent(String name, int width){
		this(name , width , 1 , 1);
	}
	
	public HeadContent(String name, int width, int rowspan, int column) {
		super();
		this.name = name;
		this.width = width;
		this.rowspan = rowspan;
		this.column = column;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public int getWidth() {
		return width;
	}

	public void setWidth(int width) {
		this.width = width;
	}

	public int getRowspan() {
		return rowspan;
	}

	public void setRowspan(int rowspan) {
		this.rowspan = rowspan;
	}

	public int getColumn() {
		return column;
	}

	public void setColumn(int column) {
		this.column = column;
	}
	
}