package com.lwb.excel.test;

import java.io.File;
import com.lwb.excel.main.AbstractCallBack;
import com.lwb.excel.main.ReadCell;
import com.lwb.excel.main.ReadExcel;

public class mainTest {

	public static void main(String[] args) {
		
		AbstractCallBack cb = new AbstractCallBack() {
			@Override
			public void afterReadLine(String[] line, int row) {
				System.out.println(row);
			}
		};
		ReadCell rc = new ReadCell() {
			@Override
			public void afterReadCell(String cell, int row, int col) {
				System.out.println(cell);
			}
		};
		cb.setReadCell(rc);
		
		ReadExcel  re = new ReadExcel(new File("C:\\Users\\liaowb\\Desktop\\GGSN_RNC.xlsx"));
		re.setCallBack(cb);
		re.setSheetNum(2);
		try {
			re.readExcel();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}
