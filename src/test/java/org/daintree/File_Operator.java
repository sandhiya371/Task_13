package org.daintree;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class File_Operator {

	public static void main(String[] args) {

		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            // Create a new sheet
            XSSFSheet sheet = workbook.createSheet("Sheet1");

            // Create column headers
            XSSFRow headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");
            headerRow.createCell(2).setCellValue("Email");

            // Data rows
            Object[][] data = {
                    {"John Doe", 30, "john@test.com"},
                    {"Jane Doe", 28, "jane@test.com"},
                    {"Bob Smith", 35, "bob@example.com"},
                    {"Swapnil", 37, "swapnil@example.com"}
            };

            // Write data to the sheet
            int rowNumber = 1;
            for (Object[] rowData : data) {
                XSSFRow row = sheet.createRow(rowNumber++);
                int cellNumber = 0;
                for (Object cellData : rowData) {
                    XSSFCell cell = row.createCell(cellNumber++);
                    if (cellData instanceof String) {
                        cell.setCellValue((String) cellData);
                    } else if (cellData instanceof Integer) {
                        cell.setCellValue((Integer) cellData);
                    }
                }
            }

            // Save the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Sandhiya\\eclipse-workspace\\FileOperator\\ExcelFile\\Task13.xlsx")) {
                workbook.write(fileOut);
                System.out.println("Excel file created successfully.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

//Output:
//	
//	Excel file created successfully.
//	(The Values of the data has been added in the Excel in the name of TASK13 in Excel folder.)

		
