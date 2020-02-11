package ExcelPractice.ExcelPractice;

import java.io.File;
import java.io.FileInputStream;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelStuff {
	// Create a utility method to store all sheetData 
    // in two dimensional String Array
    
    // method name : getAllSheetDate
    // return type : String[][]
    // params  :  FileName as String , SheetName 

	public static void main(String[] args) throws Exception {
		
		//printAllSheetData();
		String[][] result = getAllSheetData("MOCK_DATA.xlsx", "data");
		System.out.println(  Arrays.deepToString(result)  );
	}
	
	public static void printAllSheetData() throws Exception {
		File excelFile = new File("MOCK_DATA.xlsx");
		 Workbook wb = WorkbookFactory.create(excelFile);	
		 
		 Sheet sheet = wb.getSheetAt(0);
		 int rowCount = sheet.getPhysicalNumberOfRows();
		 int colCount = sheet.getRow(0).getLastCellNum();
		 
		 for(int i=0; i<rowCount; i++) {
			 System.out.println(" row number: " +i);
			 for(int j=0; j<colCount; j++) {
				 Cell cell = sheet.getRow(i).getCell(j);
				 System.out.print( cell.toString() + " | ");
			        
		      }
		      System.out.println();
				 
			 
		 }
		 wb.close();
	}
	
	public static String[][] getAllSheetData(String filePath, String SheetName) throws Exception{
		//File excelFile = new File("MOCK_DATA.xlsx");
		FileInputStream fis = new FileInputStream(filePath);
		 Workbook wb = WorkbookFactory.create(fis);	
		 
		 Sheet sheet = wb.getSheet(SheetName);
		 int rowCount = sheet.getPhysicalNumberOfRows();
		 int colCount = sheet.getRow(0).getLastCellNum();
		 
		// String[][] data = new String[11][11];
		 String[][] data = new String[rowCount][colCount];

		    for (int i = 0; i < rowCount; i++) {

		      //System.out.println(" row number : " + (i + 1));

		      for (int j = 0; j < colCount; j++) {

		        Cell cell = sheet.getRow(i).getCell(j);
		        data[i][j] = cell.toString() ; 
		        //System.out.print(cell.toString() + " | ");

		      }
		      //System.out.println();

		    }
		    fis.close();
		    wb.close();
		    
		    return data ; 
		 
	}
	
	public String getCellData(String filePath, String SheetName,int rowIndex, int colIndex) throws Exception {
		String[][] result = getAllSheetData(filePath, SheetName); 
		  return result[rowIndex][colIndex] ; 
		  
	}
}
