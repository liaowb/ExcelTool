package com.lwb.excel.main;

/**
 * 读取一行接口
 * @author liaowb
 *
 */
public interface ReadLine {

	/**
	 * 每读取完一行调用的方法
	 * @param line 该行的数据
	 */
	public void afterReadLine(String[] line,int row);
	
}
