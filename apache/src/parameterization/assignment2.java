package parameterization;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class assignment2 {

	public static void main(String[] args) throws IOException {
		String path="C:\\Users\\Kanchan\\Desktop\\project\\Book2.xlsx";
		FileInputStream File = new FileInputStream(path);
		//for workbook
		XSSFWorkbook workbook=new XSSFWorkbook(File); 
		//For sheet
		XSSFSheet sheet=workbook.getSheet("Sheet1");
		//for Row
		int Rows=sheet.getLastRowNum();
		int cells=sheet.getRow(1).getLastCellNum();
		for(int r=0; r<=Rows;r++)
		{
			XSSFRow row=sheet.getRow(r);
			for(int c=0;c<cells;c++) 
			{
				XSSFCell cell=row.getCell(c);
				switch(cell.getCellType()){
				case STRING:System.out.print(cell.getStringCellValue()); break; 
				case NUMERIC:System.out.print(cell.getNumericCellValue()); break;
				case BOOLEAN:System.out.print(cell.getBooleanCellValue()); break;
				}
				System.out.print("|");
			}
			System.out.println();
			
		}
		
		


	}

}
