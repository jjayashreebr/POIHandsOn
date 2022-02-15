package codebase;

import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFromExcel {
	public static void main(String[] args) throws IOException {
		String path = System.getProperty("user.dir") + "/Writesheet.xlsx";
		File excel = new File(path);
		FileInputStream fis = new FileInputStream(excel);

		@SuppressWarnings("resource")
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet ws = wb.getSheet("Employee Info");

		int rowNum = ws.getLastRowNum() + 1;
		int colNum = ws.getRow(0).getLastCellNum();

		for (int i = 0; i < rowNum; i++) {
			XSSFRow row = ws.getRow(i);
			for (int j = 0; j < colNum; j++) {
				Cell cell = row.getCell(j);
				
				 switch (cell.getCellType()) {
	               case NUMERIC:
	                  System.out.print(cell.getNumericCellValue() + " \t\t ");
	                  break;
	               
	               case STRING:
	                  System.out.print(
	                  cell.getStringCellValue() + " \t\t ");
	                  break;
				default:
					  System.out.print(
			                  cell.getStringCellValue() + " \t\t ");
					break;
	            }
	         }
       System.out.println();
		}

		}
		
}

