package Assignment3;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readdata3 {
public static void main(String[] args) throws IOException {
	
		
		String path ="C:\\Users\\Kanchan\\Desktop\\project\\Employee.xlsx";
		
		XSSFWorkbook workbook  =new XSSFWorkbook(path); 
		
		XSSFSheet Sheet =workbook.getSheet("sheet1"); 
		
		int RowCount  = Sheet.getLastRowNum(); 
		
		int CellCount =Sheet.getRow(0).getLastCellNum();
		
		for(int r=0;r<=RowCount;r++) {  
			
			XSSFRow CurrentRow =Sheet.getRow(r);
			
			for(int c=0;c<CellCount ;c++) {
				
			String value = CurrentRow.getCell(c).toString();
			
			System.out.print(" | "+value);
			}
			
			System.out.println();
		}
		
		
	}
}
