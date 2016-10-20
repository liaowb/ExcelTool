package com.lwb.excel.main;

/**
 * 读取单元格接口
 * @author liaowb
 *
 */
public interface ReadCell {

	/**
	 * 每读取完一个单元格调用的方法
	 * @param cell 单元格的值
	 */
	public void afterReadCell(String cell,int row,int col);
	
}
