package Assignment3;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readdta4 {
	
    static  Object CellType;
	public static void main(String[] args) throws IOException {
		String path="C:\\Users\\Kanchan\\Desktop\\project\\Employee.xlsx";
		FileInputStream File = new FileInputStream(path);
		//for workbook
		XSSFWorkbook workbook=new XSSFWorkbook(File); 
		//For sheet
		XSSFSheet sheet=workbook.getSheet("Sheet1");
		//for Row
		int Row=sheet.getLastRowNum();

		int Col=sheet.getRow(1).getLastCellNum();
		
		for(int r=0;r<=Row;r++)
		 {
			 XSSFRow row = sheet.getRow(r);
			 {
				 for(int c=0;c<Col;c++)
				 {
				  	 XSSFCell cell = row.getCell(c);
				  	 
			 	  
					if(cell.getCellType().equals(CellType.toString()))
				  	 {
				  		  System.out.print(cell.getStringCellValue());
				  	 }
			 	  	 else if(cell.getCellType().equals(CellType.toString()))
			 	  	 {
			 	  		System.out.print(cell.getNumericCellValue());
			 	  	 }
			 	  	 else if(cell.getCellType().equals(CellType.toString()))
			 	  	 {
			 	  		System.out.print(cell.getBooleanCellValue());
			 	  	 }
			 	  	System.out.print(" | ");
				 }
				 
			 }
			 System.out.println();
		 }
	}
}
