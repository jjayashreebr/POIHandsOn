package codebase;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.util.Date;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Apache2 {
	public static void main(String[] args) throws Exception {
		@SuppressWarnings("resource")
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet("cell types");

		XSSFRow row = spreadsheet.createRow((short) 2);
		row.createCell(0).setCellValue("Type of Cell");
		row.createCell(1).setCellValue("cell value");
		row.createCell(0).setCellValue("set cell type BOOLEAN");
		row.createCell(1).setCellValue(true);

		row = spreadsheet.createRow((short) 5);
		row.createCell(0).setCellValue("set cell type date");
		row.createCell(1).setCellValue( LocalDate.now());

		row = spreadsheet.createRow((short) 6);
		row.createCell(0).setCellValue("set cell type numeric");
		row.createCell(1).setCellValue(20);
		row = spreadsheet.createRow((short) 7);
		row.createCell(0).setCellValue("set cell type string");
		row.createCell(1).setCellValue("A String");
		String path = System.getProperty("user.dir") + "/sheetcelltypes.xlsx";
		FileOutputStream out = new FileOutputStream(new File(path));
		workbook.write(out);
		out.close();
		System.out.println("typesofcells.xlsx written successfully");
	}
}
