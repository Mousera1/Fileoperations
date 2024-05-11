package fileoperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelDemo {

	public static void main(String[] args) throws IOException {
		File src = new File("D:\\guvi\\Excelworksheetdemo.xlsx"); // specify location of excel file

		FileInputStream fis = new FileInputStream(src); // load file

		XSSFWorkbook wb = new XSSFWorkbook(fis); // load workbook

		XSSFSheet sh = wb.getSheet("Demosheet"); // load worksheet

		System.out.println(sh.getSheetName()); // print the name of loaded sheet

		System.out.println(sh.getRow(0).getCell(0).getStringCellValue()); // print Username from excel sheet

		System.out.println(sh.getRow(2).getCell(1).getStringCellValue()); // print p2 from excel sheet

		System.out.println("Total rows:" + sh.getPhysicalNumberOfRows()); // print total no of rows
		System.out.println("Total Columns:" + sh.getRow(0).getPhysicalNumberOfCells()); // print total no of columns

	}

}
