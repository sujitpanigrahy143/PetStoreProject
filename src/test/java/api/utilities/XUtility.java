package api.utilities;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XUtility {
	
	public FileInputStream fi;
	public FileOutputStream fo;
	public XSSFWorkbook workBook;
	public XSSFSheet sheet;
	public XSSFRow rows;
	public XSSFCell cell;
	public CellStyle style;
	String path;
	String data;
	
	public XUtility(String path) {
		this.path = path;
	}
	
	public int getRowCount(String sheetName) throws IOException {
		
		fi = new FileInputStream(path);
		workBook = new XSSFWorkbook(fi);
		sheet = workBook.getSheet(sheetName);
		int rowCount = sheet.getLastRowNum();
		workBook.close();
		return rowCount;
	}

	public int getCellCount(String sheetName,int rowNum) throws IOException {
			
		fi = new FileInputStream(path);
		workBook = new XSSFWorkbook(fi);
		sheet = workBook.getSheet(sheetName);
		rows = sheet.getRow(rowNum);
		int cellCount = rows.getLastCellNum();
		workBook.close();
		return cellCount;
	}
	
	public String getCellData(String sheetName,int rowNum,int colNum) throws IOException {
		
		fi = new FileInputStream(path);
		workBook = new XSSFWorkbook(fi);
		sheet = workBook.getSheet(sheetName);
		rows = sheet.getRow(rowNum);
		cell = rows.getCell(colNum);
		DataFormatter formatter = new DataFormatter();
		
		try {
			data = formatter.formatCellValue(cell);
		}
		catch(Exception ex) {
			data="";
		}
		
		workBook.close();
		fi.close();
		return data;
	}
	public String setCellData(String sheetName,int rowNum,int colNum,String data) throws IOException {
		
		File xFile = new File(path);
		if(!xFile.exists()) {
			workBook = new XSSFWorkbook() ;
			fo = new FileOutputStream(path);
			workBook.write(fo);
		}
		///ITS NOT COMPLETED
		//fi = new File InputStream(path);
		workBook = new XSSFWorkbook(fi);
		sheet = workBook.getSheet(sheetName);
		rows = sheet.getRow(rowNum);
		cell = rows.getCell(colNum);
		DataFormatter formatter = new DataFormatter();
		//String data;
		try {
			data = formatter.formatCellValue(cell);
		}
		catch(Exception ex) {
			data="";
		}
		
		workBook.close();
		fi.close();
		return data;
		}
}
