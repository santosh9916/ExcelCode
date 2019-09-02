package Property_Excel_relatedCode;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromPropertyAndWriteIntoExcel {
	Properties prop;
	String prop_file = "D:\\workspaceFinal\\Selenium3_Latest\\src\\Property_Excel_relatedCode\\config.properties";
	String excel_file = "D:\\workspaceFinal\\Selenium3_Latest\\src\\Property_Excel_relatedCode\\Testdata.xlsx";
	FileInputStream fis_prop;
	FileOutputStream fis_excelWrite;

	public void m1() throws IOException, EncryptedDocumentException, InvalidFormatException {
		
		// Load properties file data
		prop = new Properties();
		fis_prop = new FileInputStream(new File(prop_file));
		prop.load(fis_prop);
		
		// Excel code
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.getSheetAt(1);

		sheet.createRow(0).createCell(0).setCellValue(prop.getProperty("username"));
		fis_excelWrite = new FileOutputStream(new File(excel_file));
		workbook.write(fis_excelWrite);
		fis_excelWrite.close();
		System.out.println("gfgcontribute.xlsx written successfully on disk.");

	}

	public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {
		ReadDataFromPropertyAndWriteIntoExcel readwrite = new ReadDataFromPropertyAndWriteIntoExcel();
		readwrite.m1();
	}

}
