package fileoperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteonExcelTask {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook wb  = new XSSFWorkbook() ; // create object of workbook
		XSSFSheet sh = wb.createSheet("Person Details"); // create sheet
		
		ArrayList<Object[]> empData = new ArrayList<>(); // create arraylist
		empData.add(new Object[] {"Name", "Age","Email"});
		empData.add(new Object[] {"John Doe", 30,"john@test.com"});
		empData.add(new Object[] {"John Doe", 28,"john@test.com"});
		empData.add(new Object[] {"Bob Smith", 35,"jacky@example.com"});
		empData.add(new Object[] {"Swapnil", 37,"swapnil@example.com"});
		
		
		int rownum = 0;
		
		//outer loop for rows
		for (Object [] emp:empData) {
			XSSFRow row = sh.createRow(rownum++); // for rows
			int cellnum = 0;
			//inner loop for columns
			for(Object value:emp) {
				XSSFCell cell = row.createCell(cellnum++);
				if(value instanceof String)
					cell.setCellValue((String)value);
				if(value instanceof Integer)
					cell.setCellValue((Integer)value);
				if(value instanceof Boolean)
					cell.setCellValue((Boolean)value);
				
			}
		}
			
			String filePath = ".\\datafiles\\Persondetails.xlsx"; // giving filepath
			FileOutputStream fos = new FileOutputStream(filePath); // create object of fileoutputstream
			
			wb.write(fos);
			fos.close(); // close 
			
			System.out.println("Persondetails.xlsx file writtern successfully");
			
		

	}

}
