package fileoperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateSheetExcel {

    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook(); // Create a new XSSFWorkbook (for .xlsx format)
        String filePath = "./datafiles/workbook_with_sheet1.xlsx"; // File path for the new Excel file
        String sheetName = "Sheet1"; // Name for the new sheet

        FileOutputStream fos = new FileOutputStream(filePath) ;
            // Create a new sheet with the specified name
            Sheet sheet = workbook.createSheet(sheetName);

            // Write the workbook to the output stream
            workbook.write(fos);
            fos.close();
            System.out.println("Workbook with sheet '" + sheetName + "' created successfully");
         
        }}