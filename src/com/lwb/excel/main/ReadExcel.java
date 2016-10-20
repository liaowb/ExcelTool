package com.lwb.excel.main;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.lwb.excel.util.Common;
import com.lwb.excel.util.Util;

public class ReadExcel {

	private File excelFile = null;
	private CallBack callBack = null;
	private Integer sheetNum = null;
	
	
	public ReadExcel(){};
	public ReadExcel(File excelFile) {
		this.excelFile = excelFile;
	}
	public ReadExcel(File excelFile, CallBack callBack) {
		this.excelFile = excelFile;
		setCallBack(callBack);
	}
	
	
	/**
     * read the Excel file
     * @param path the path of the Excel file
     * @return
     * @throws IOException
     */
    public void readExcel() throws IOException {
    	String path = excelFile.getAbsolutePath();
        if (path == null || Common.EMPTY.equals(path)) {
            return;
        } else {
            String postfix = Util.getPostfix(path);
            if (!Common.EMPTY.equals(postfix)) {
                if (Common.OFFICE_EXCEL_2003_POSTFIX.equals(postfix)) {
                    readXls(path);
                } else if (Common.OFFICE_EXCEL_2010_POSTFIX.equals(postfix)) {
                    readXlsx(path);
                }
            } else {
                System.out.println(path + Common.NOT_EXCEL_FILE);
            }
        }
        return;
    }

    /**
     * Read the Excel 2010
     * @param path the path of the excel file
     * @return
     * @throws IOException
     */
    private void readXlsx(String path) throws IOException {
        System.out.println(Common.PROCESSING + path);
        InputStream is = new FileInputStream(path);
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
        read(xssfWorkbook);
    }

    /**
     * Read the Excel 2003-2007
     * @param path the path of the Excel
     * @return
     * @throws IOException
     */
    private void readXls(String path) throws IOException {
        System.out.println(Common.PROCESSING + path);
        InputStream is = new FileInputStream(path);
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
        read(hssfWorkbook);
    }

    /**
     * 获取单元格的值
     * @param Cell
     * @return
     */
    @SuppressWarnings("static-access")
	private String getValue(Cell Cell) {
    	if(Cell == null){
    		return "";
    	}
        if (Cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(Cell.getBooleanCellValue());
        } else if (Cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            return String.valueOf(Cell.getNumericCellValue());
        } else if( Cell.getCellType() == Cell.CELL_TYPE_STRING ){
            return String.valueOf(Cell.getStringCellValue());
        }else{
        	Cell.setCellType(org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING); 
        	return String.valueOf(Cell.getStringCellValue());
        }
    }
    
    /**
     * 读取工作簿
     * @param workbook
     */
	private void read(Workbook workbook){
		if( this.sheetNum != null ){
			readSheet(workbook, sheetNum);
		}else{
			for (int SheetNum = 0; SheetNum < workbook.getNumberOfSheets(); SheetNum++) {
	    		readSheet(workbook, SheetNum);
	    	}
		}
    }
	
	public void readSheet( Workbook workbook ,Integer sheetNum ){
		Sheet sheet = workbook.getSheetAt(sheetNum);
        if (sheet == null) {
            return;
        }
        int rowCount = sheet.getLastRowNum();
        for (int rowNum = 0; rowNum <= rowCount; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
            	int colCount = row.getLastCellNum();
            	if(colCount < 0){
            		continue;
            	}
            	String[] line = new String[colCount];
            	for(int colNum = 0 ; colNum < colCount ; colNum ++){
            		Cell cell = row.getCell(colNum);
            		String strcell = getValue(cell);
            		line[colNum] = strcell;
            		this.callBack.afterReadCell(strcell, rowNum, colNum);
            	}
            	this.callBack.afterReadLine(line, rowNum);
            }
        }
        this.callBack.afterReadSheet(sheet, sheetNum);
	}
    
	public void setExcelFile(File excelFile) {
		this.excelFile = excelFile;
	}
	
	/**
	 * 设置回调函数
	 * @param callBack
	 */
	public void setCallBack(CallBack callBack) {
		this.callBack = callBack;
	}
	
	/**
	 * 设置读Excel文件的页码
	 * @param sheetNum sheetNum >= 1
	 */
	public void setSheetNum(Integer sheetNum ){
		this.sheetNum = sheetNum;
	}
}
