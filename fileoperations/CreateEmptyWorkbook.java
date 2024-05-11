package fileoperations;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class CreateEmptyWorkbook {
    public static void main(String[] args) {
        // Create a new workbook
        Workbook workbook = new XSSFWorkbook(); // For .xlsx format, use XSSFWorkbook

        // Write the workbook to a file
        try (FileOutputStream outputStream = new FileOutputStream("emptyWorkbook1.xlsx")) 
        {
            workbook.write(outputStream);
            System.out.println("Empty Excel workbook created successfully.");
        } 
        catch (IOException e) {
            System.out.println("Error occurred while creating empty Excel workbook: " + e.getMessage());
        } 
        finally 
        {
            // Close the workbook
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}