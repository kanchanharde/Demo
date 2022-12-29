package parameterization;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Readdata {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		String path="C:\\Users\\Kanchan\\Desktop\\project\\sample.xlsx";
		FileInputStream file=new FileInputStream(path);
		String Data=WorkbookFactory.create(file).getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue();
        System.out.println(Data);
	}

}
