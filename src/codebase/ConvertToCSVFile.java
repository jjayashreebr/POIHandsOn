package codebase;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.opencsv.CSVWriter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

public class ConvertToCSVFile {

	public static void main(String[] args) throws IOException {
		
		
		String path = System.getProperty("user.dir") + "/Fruits.xlsx";
		try (// create a new XLSX file
		XSSFWorkbook workbook1 = new XSSFWorkbook()) {
			OutputStream outputStream = new FileOutputStream(path);
			// create a new sheet
			XSSFSheet sheet = workbook1.createSheet("fruits");
			// create a new sheet
			
			for(int i=0;i<3;i++) {
				Row row  = sheet.createRow(i);
			for(int j=0;j<3;j++) {
			
			Cell cell1 = row.createCell(j);
			// set cell value
			cell1.setCellValue("data"+i);
			}}
			// save file
			workbook1.write(outputStream);
			outputStream.close();
		}
		// Open and existing XLSX file
		FileInputStream fileInStream = new FileInputStream(path);
		XSSFWorkbook workBook = new XSSFWorkbook(fileInStream);
		XSSFSheet selSheet = workBook.getSheetAt(0);
		// Loop through all the rows
		Iterator<?> rowIterator = selSheet.iterator();
		//csv file
		path = System.getProperty("user.dir") + "/output.csv";
		CSVWriter csvWriter = new CSVWriter(new FileWriter(path));
		
		while (rowIterator.hasNext()) {
		  Row row = (Row) rowIterator.next();
	
		  String output[] = new String[3];
		  for (int j = 0; j < 3 ; j++){
		  Cell cell = (Cell) row.getCell(j);
		  output[j]=cell.getStringCellValue();
	       }
		  csvWriter.writeNext(output);
		}
		csvWriter.close();
		workBook.close();
		fileInStream.close();
		
		
		
		
		
		
	}

}
