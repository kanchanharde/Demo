package parameterization;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readdata2 {

		public static void main(String[] args) throws EncryptedDocumentException, IOException {
			String path="C:\\Users\\Kanchan\\Desktop\\project\\sample.xlsx";
			FileInputStream file=new FileInputStream(path);
			XSSFWorkbook Workbook=new XSSFWorkbook(file);
			String Data= Workbook.getSheetAt(0).getRow(1).getCell(0).getStringCellValue();
			double Data1= Workbook.getSheetAt(0).getRow(1).getCell(1).getNumericCellValue();
	        System.out.println(Data);
	        System.out.println(Data1);
		}

	}


