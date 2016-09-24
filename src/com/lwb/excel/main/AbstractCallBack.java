package com.lwb.excel.main;

import org.apache.poi.ss.usermodel.Sheet;

public abstract class AbstractCallBack implements CallBack {
	
	private ReadCell readCell = null;
	private ReadLine readLine = null;
	private ReadSheet readSheet = null;
	
	public ReadCell getReadCell() {
		return readCell;
	}
	public void setReadCell(ReadCell readCell) {
		this.readCell = readCell;
	}
	public ReadLine getReadLine() {
		return readLine;
	}
	public void setReadLine(ReadLine readLine) {
		this.readLine = readLine;
	}
	public ReadSheet getReadSheet() {
		return readSheet;
	}
	public void setReadSheet(ReadSheet readSheet) {
		this.readSheet = readSheet;
	}
	
	
	@Override
	public void afterReadCell(String cell, int row, int col) {
		if( readCell != null )
			readCell.afterReadCell(cell, row, col);
	}
	@Override
	public void afterReadLine(String[] line, int row) {
		if( readLine != null )
			readLine.afterReadLine(line, row);
	}
	@Override
	public void afterReadSheet(Sheet sheet, int num) {
		if( readSheet != null ) 
			readSheet.afterReadSheet(sheet, num);
	}
	
}
