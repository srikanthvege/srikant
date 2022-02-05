package com.test;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.aspose.cells.LoadFormat;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;

public class Test {

	public static String originalFilePath = "C:\\Users\\suvarchala.vege\\Desktop\\password.xlsx";
	
	public static void main(String[] args) throws Exception {
		convertToCSV();
		decryptPassword();
	}
	
	public static void convertToCSV() throws Exception{
		Path source = Paths.get(originalFilePath);
		Path target = Paths.get("C:\\Users\\suvarchala.vege\\Desktop\\passwordCSV.csv");
		Files.copy(source, target);
	}
	public static void decryptPassword() throws Exception {
		LoadOptions loadOptions = new LoadOptions(LoadFormat.XLSX);
		loadOptions.setPassword("abcd");
		Workbook workbook = new Workbook(originalFilePath, loadOptions);
		workbook.getSettings().setPassword(null);
		workbook.save("C:\\Users\\suvarchala.vege\\Desktop\\decrypted-workbook.xlsx");
		//remove("C:\\Users\\suvarchala.vege\\Desktop\\decrypted-workbook.xlsx");
	}
	
	public static void remove(String filePath) throws Exception{
		FileInputStream inputStream = new FileInputStream(new File(filePath));
		  XSSFWorkbook workBook = new XSSFWorkbook(inputStream);
		  workBook.removeSheetAt(1);
		  FileOutputStream outFile =new FileOutputStream(new File(filePath));
		  workBook.write(outFile);
		  outFile.close();
	}
}
