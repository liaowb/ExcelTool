package com.lwb.excel.main;

import org.apache.poi.ss.usermodel.Sheet;

public interface ReadSheet {

	/**
	 * 每读取完一页调用的方法
	 * @param sheet 该页对象
	 */
	public void afterReadSheet(Sheet sheet,int num);
	
}
