package parameterization;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readdata3 {
public static void main(String[] args) throws IOException {
	
		
	    String Path = "C:\\Users\\Kanchan\\Desktop\\project\\Book2.xlsx";
		
		FileInputStream File = new FileInputStream(Path);
		
		XSSFWorkbook Workbook =new XSSFWorkbook(File);
		
		XSSFSheet Sheet =Workbook.getSheet("Sheet1");
		
		int Row = Sheet.getLastRowNum();
		System.out.println(Row);
		
	    int Col	=Sheet.getRow(0).getLastCellNum();
	    System.out.println(Col);
	    
	    for(int r=0;r<=Row;r++) {  //Row
	    	
	    	XSSFRow row =Sheet.getRow(r);
	    	
	    	for(int C=0;C<Col;C++) {
	    		
	    		XSSFCell cell =row.getCell(C);
	    		
	    		switch(cell.getCellType()) {
	    		
	    		case STRING :System.out.print(cell.getStringCellValue()); 
	    		break;
	    		
	    		case NUMERIC :System.out.print(cell.getNumericCellValue());
	    		break;
	    		
	    		case BOOLEAN : System.out.print(cell.getBooleanCellValue());
	    		break;
				default:
					break;
	    		}
	    		System.out.print(" | ");
	    	}
	    	System.out.println();
	    }
		
	}


}
