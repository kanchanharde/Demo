package Assignment3;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Readdata {
	public static void main(String[] args) throws EncryptedDocumentException, IOException  {
		String path="C:\\Users\\Kanchan\\Desktop\\Project\\Employee.xlsx";
		FileInputStream File=new FileInputStream (path);
		String data=WorkbookFactory.create(File).getSheet("Sheet1").getRow(3).getCell(3).getStringCellValue();
        System.out.println(data);
	}

}
