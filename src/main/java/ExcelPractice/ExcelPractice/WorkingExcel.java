package ExcelPractice.ExcelPractice;

import java.io.File;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class WorkingExcel {

	public static void main(String[] args) throws Exception{
		
		//Workbook --> Sheet ---> Row--> Cell 
		File excelFile = new File("MOCK_DATA.xlsx") ; 
		Workbook wb = WorkbookFactory.create(excelFile);
		
		System.out.println(wb.getNumberOfSheets() );   
		
		//Sheet sh = wb.getSheet("data");
		Sheet sh = wb.getSheetAt(0); 
		Row row1 = sh.getRow(0) ; 
		Cell c1 =  row1.getCell(1) ; 
		System.out.println( c1 );

		int columnCount= row1.getLastCellNum();
		System.out.println(columnCount);
		
		int  rowCount = sh.getLastRowNum();
		System.out.println(rowCount);
		
//		int columnCountInFirstRow = row1.getFirstCellNum();
//		System.out.println(columnCountInFirstRow);
		wb.close();


	}

}
