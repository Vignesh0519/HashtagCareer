package Testcase;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class appPage {
	private static WebDriver driver;
	private static String excelFilePath = "./test-data/Testdata.xlsx";
	private static String url = "https://www.hashtag-ca.com/careers/apply?jobCode=QAE001";

	public static void main(String[] args) {
		setup();
		readExcelAndApply();
		tearDown();
	}

	
	public static void setup() {
		// Set up WebDriver and open the browser
		System.setProperty("webdriver.chrome.driver", "F:\\Servers's & Connectors\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}


	public static void readExcelAndApply() {
		try {
			// Create a FileInputStream object for the Excel file
			FileInputStream fis = new FileInputStream(excelFilePath);

			// Create a Workbook object from the Excel file
			XSSFWorkbook workbook = new XSSFWorkbook(fis);

			// Access the desired sheet in the Excel file (by index or name)
			XSSFSheet sheet = workbook.getSheetAt(0); // Assuming the first sheet

			// Navigate to the URL
			driver.get(url);
			// Iterate through the rows in the sheet
			for (Row row : sheet) {
				// Extract data from Excel columns
				
				String name = row.getCell(0).getStringCellValue();
				String email = row.getCell(1).getStringCellValue();
				String phone = row.getCell(2).getStringCellValue();
				String resumeFilePath = row.getCell(3).getStringCellValue();
				String description = row.getCell(4).getStringCellValue();

				// Fill in the form fields on the web page
				WebElement nameField = driver.findElement(By.name("name"));
				nameField.sendKeys(name);

				WebElement emailField = driver.findElement(By.name("email"));
				emailField.sendKeys(email);

				WebElement phoneField = driver.findElement(By.name("phone"));
				phoneField.sendKeys(phone);

				WebElement resumeField = driver.findElement(By.xpath("//*[@id=\"inputFile\"]"));
				resumeField.sendKeys(resumeFilePath);

				WebElement descriptionField = driver.findElement(By.name("description"));
				descriptionField.sendKeys(description);

				// Click on the Apply Now button
				WebElement applyNowButton = driver.findElement(By.xpath("//*[@id=\"contact-form\"]/div/div[7]/div/button[1]"));
				applyNowButton.click();

				// Wait for confirmation or success message, handle as needed
				// You can use explicit or implicit waits here
				// Example: Thread.sleep(2000); or WebDriverWait

				// Clear the form fields for the next iteration
				nameField.clear();
				emailField.clear();
				phoneField.clear();
				resumeField.clear();
				descriptionField.clear();
			}

			// Close the workbook and FileInputStream
			workbook.close();
			fis.close();

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
		 public static void tearDown() {
		        // Close the browser
		        driver.quit();
	}

}
