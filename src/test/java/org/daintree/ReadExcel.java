package org.daintree;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

public class ReadExcel {
    public static void main(String[] args) {
        try (FileInputStream file = new FileInputStream("C:\\Users\\Sandhiya\\eclipse-workspace\\FileOperator\\ExcelFile\\Task13.xlsx")) {
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0); // assuming you want to read the first sheet

            for (Row row : sheet) {
                for (Cell cell : row) {
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println(); // Move to the next line after each row
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
				
    }
    
    }


//Output:
//	
//	Name			Age		Email	
//	John Doe		30.0	john@test.com	
//	Jane Doe		28.0	jane@test.com	
//	Bob Smith		35.0	bob@example.com	
//	Swapnil			37.0	swapnil@example.com