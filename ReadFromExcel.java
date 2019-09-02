package Property_Excel_relatedCode;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class ReadFromExcel {
	private static WebDriver driver;
	String excel_file = "D:\\workspaceFinal\\Selenium3_Latest\\src\\Property_Excel_relatedCode\\Testdata2.xlsx";
	FileInputStream file;
	XSSFWorkbook workbook;
	XSSFSheet sheet;

	public void init() {
		System.setProperty("webdriver.chrome.driver", "D:\\Library\\chrome_76\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().timeouts().pageLoadTimeout(10, TimeUnit.SECONDS);
		driver.get("https://www.google.com");

	}

	public void enhanceForLoop() {
		try {
			file = new FileInputStream(new File(excel_file));
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheet("Links");

			for(Row row : sheet){
			 for(Cell cell : row){
				 switch (cell.getCellType()){
				 case Cell.CELL_TYPE_NUMERIC:
					 System.out.print(cell.getNumericCellValue());
					 break;
				 case Cell.CELL_TYPE_STRING:
					 System.out.print(cell.getStringCellValue());
					 break;
				 }
			 }
			 System.out.println("");
			}
			file.close();
			
		} catch (Exception e) {
			
			e.printStackTrace();
		}
	}

	public void read_iterator() {
		try {
			file = new FileInputStream(new File(excel_file));

			// Create Workbook instance holding reference to .xlsx file
			workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			sheet = workbook.getSheetAt(2);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print(cell.getNumericCellValue());
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print(cell.getStringCellValue());
						break;
					}
				}
				System.out.println("");
			}
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		ReadFromExcel rd = new ReadFromExcel();
		// rd.read_iterator();
		rd.enhanceForLoop();

	}

}
