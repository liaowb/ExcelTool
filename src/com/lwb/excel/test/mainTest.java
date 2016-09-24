package com.lwb.excel.test;

import java.io.File;

import org.apache.poi.ss.usermodel.Sheet;

import com.lwb.excel.main.AbstractCallBack;
import com.lwb.excel.main.ReadCell;
import com.lwb.excel.main.ReadExcel;
import com.lwb.excel.main.ReadSheet;

public class mainTest {

	public static void main(String[] args) {
		
		AbstractCallBack cb = new AbstractCallBack() {
			@Override
			public void afterReadLine(String[] line, int row) {
				System.out.println(row);
			}
			
			@Override
			public void setReadSheet(ReadSheet readSheet) {
				super.setReadSheet(readSheet);
				readSheet = new ReadSheet() {
					@Override
					public void afterReadSheet(Sheet sheet, int num) {
						System.out.println("1232321232132");
					}
				};
				super.setReadSheet(readSheet);
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
