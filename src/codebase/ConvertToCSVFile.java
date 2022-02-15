package codebase;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;

public class ConvertToCSVFile {

	public static void main(String[] args) throws IOException {
		
		
		String path = System.getProperty("user.dir") + "/LargeDocument.xlsx";
		try (// create a new XLSX file
		XSSFWorkbook workbook1 = new XSSFWorkbook()) {
			OutputStream outputStream = new FileOutputStream(path);
			// create a new sheet
			XSSFSheet sheet = workbook1.createSheet("Apache POI XSSF");
			// create a new sheet
			Row row1  = sheet.createRow(1);
			// create a new cell
			Cell cell1 = row1.createCell(1);
			// set cell value
			cell1.setCellValue("File Format Developer Guide");
			
			Cell cell2 = row1.createCell(2);
			cell2.setCellValue("File Format Developer Guide");
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
		while (rowIterator.hasNext()) {
		  Row row = (Row) rowIterator.next();
		  // Loop through all rows and add","
		  Iterator<?> cellIterator = row.cellIterator();
		  StringBuffer stringBuffer = new StringBuffer();
		  while (cellIterator.hasNext()) {
		  Cell cell = (Cell) cellIterator.next();
		  if (stringBuffer.length() != 0) {
		    stringBuffer.append(",");
		  }
		  stringBuffer.append(cell.getStringCellValue());
		  }
		  System.out.println(stringBuffer.toString());
		}
		workBook.close();
		fileInStream.close();
	}

}
