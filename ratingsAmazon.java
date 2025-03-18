package Shopping;

import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;

import CommonUtility.BlinkitId;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ratingsAmazon {

	public static void main(String[] args) throws Exception {
		System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
		WebDriver driver = new ChromeDriver();

		int count = 0;
		// int finalSp;
		
		String newName = null;
		String rating = null;
		String reviewCount = null; 
		try {
			// Read URLs from Excel file
			String filePath = ".\\input-data\\Rating input data.xlsx";
			FileInputStream file = new FileInputStream(filePath);
			Workbook urlsWorkbook = new XSSFWorkbook(file);
			Sheet urlsSheet = urlsWorkbook.getSheet("Amazon");
			int rowCount = urlsSheet.getPhysicalNumberOfRows();

			List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
					InputSize = new ArrayList<>(), uRL = new ArrayList<>(),
					UOM = new ArrayList<>();

			// Extract URLs from Excel
			for (int i = 0; i < rowCount; i++) {
				Row row = urlsSheet.getRow(i);

				if (i == 0) {
					continue;
				}

				Cell inputPidCell = row.getCell(0);
				Cell inputCityCell = row.getCell(1);
				Cell inputNameCell = row.getCell(2);
				Cell inputSizeCell = row.getCell(3);
				Cell urlCell = row.getCell(4);
				Cell uomCell = row.getCell(5);
				
				if (urlCell != null && urlCell.getCellType() == CellType.STRING) {
					String url = urlCell.getStringCellValue();
					String id = (inputPidCell != null && inputPidCell.getCellType() == CellType.STRING)
							? inputPidCell.getStringCellValue()
							: "";
					String city = (inputCityCell != null && inputCityCell.getCellType() == CellType.STRING)
							? inputCityCell.getStringCellValue()
							: "";
					String name = (inputNameCell != null && inputNameCell.getCellType() == CellType.STRING)
							? inputNameCell.getStringCellValue()
							: "";
					String size = (inputSizeCell != null && inputSizeCell.getCellType() == CellType.STRING)
							? inputSizeCell.getStringCellValue()
							: "";
					String uom = (uomCell != null && uomCell.getCellType() == CellType.STRING)
							? uomCell.getStringCellValue()
							: "";
					
					inputPid.add(id);
					InputCity.add(city);
					InputName.add(name);
					InputSize.add(size);
					uRL.add(url);
					UOM.add(uom);
					
					/*
					 * uRL.add(url); ids.add(id); offers.add(offer);
					 */

				}
			}
			// Create Excel workbook for storing results
			Workbook resultsWorkbook = new XSSFWorkbook();
			Sheet resultsSheet = resultsWorkbook.createSheet("Results");

			Row headerRow = resultsSheet.createRow(0);

			headerRow.createCell(0).setCellValue("InputPid");
			headerRow.createCell(1).setCellValue("InputCity");
			headerRow.createCell(2).setCellValue("InputName");
			headerRow.createCell(3).setCellValue("InputSize");
			headerRow.createCell(4).setCellValue("URL");
			headerRow.createCell(5).setCellValue("Name");
			headerRow.createCell(6).setCellValue("UOM");
			headerRow.createCell(7).setCellValue("Ratings");
			headerRow.createCell(8).setCellValue("Review Count");
			headerRow.createCell(9).setCellValue("Commands");
			headerRow.createCell(10).setCellValue("Remarks");
			headerRow.createCell(11).setCellValue("Correctness");
			headerRow.createCell(12).setCellValue("Percentage");
			headerRow.createCell(13).setCellValue("Name");
			

			int rowIndex = 1;

			int headercount = 0;

			for (int i = 0; i < uRL.size(); i++) {
				String id = inputPid.get(i);
				String city = InputCity.get(i);
				String name = InputName.get(i);
				String size = InputSize.get(i);
				String url = uRL.get(i);
				String uom = UOM.get(i);

				try {

					if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
						// Set "NA" values in all three columns
						Row resultRow = resultsSheet.createRow(rowIndex++);
						resultRow.createCell(0).setCellValue(id);
						resultRow.createCell(1).setCellValue(city);
						resultRow.createCell(2).setCellValue(name);
						resultRow.createCell(3).setCellValue(size);
						resultRow.createCell(4).setCellValue(url);
						resultRow.createCell(5).setCellValue("NA");
						resultRow.createCell(6).setCellValue("NA");
						resultRow.createCell(7).setCellValue("NA");
						resultRow.createCell(8).setCellValue("NA");
						resultRow.createCell(9).setCellValue("NA");
						resultRow.createCell(10).setCellValue("NA");

						System.out.println("Skipped processing for URL: " + url);
						continue; // Skip to the next iteration
					}

					driver.get(url);
					driver.manage().window().maximize();

					
					try {

						WebElement nameElement = driver.findElement(By.id("productTitle"));
						newName = nameElement.getText();
						System.out.println(newName);
					}

					catch (NoSuchElementException e) {

						WebElement nameElement = driver.findElement(By.xpath("//*[@id=\"productTitle\"]"));
						newName = nameElement.getText();
						System.out.println(newName);

					}
					System.out.println("headercount = " + headercount);

					headercount++;

					//rating
					try{
						
						WebElement ratings = driver.findElement(By.xpath("//div[@id='cm_cr_dp_d_rating_histogram']//span[@class='a-size-base a-nowrap']//span"));
						String rating1 = ratings.getText();
						rating = rating1.split(" ")[0];
						System.out.println(rating);
					}
					catch(Exception m) {
						rating = "NA";
					}
					
					
					//rating count
					
					try{
						WebElement reviewCounts = driver.findElement(By.xpath("//div[@id='cm_cr_dp_d_rating_histogram']//div[@class='a-row a-spacing-medium averageStarRatingNumerical']//span"));
						String reviewCount1 = reviewCounts.getText();
						reviewCount = reviewCount1.split(" ")[0];
						System.out.println(reviewCount);
					}
					catch(Exception y) {
						reviewCount = "NA";
					}
					
					
					Row resultRow = resultsSheet.createRow(rowIndex++);

					resultRow.createCell(0).setCellValue(id);
					resultRow.createCell(1).setCellValue(city);
					resultRow.createCell(2).setCellValue(name);
					resultRow.createCell(3).setCellValue(size);
					resultRow.createCell(4).setCellValue(url);
					resultRow.createCell(5).setCellValue(newName);
					resultRow.createCell(6).setCellValue(uom);
					resultRow.createCell(7).setCellValue(rating);
					resultRow.createCell(8).setCellValue(reviewCount);
					resultRow.createCell(9).setCellValue(" ");
					resultRow.createCell(10).setCellValue(" ");
					resultRow.createCell(11).setCellValue(" ");
					resultRow.createCell(12).setCellValue(" ");
					resultRow.createCell(13).setCellValue(" ");
					//resultRow.createCell(14).setCellValue("NA");

					System.out.println("Data extracted for URL: " + url);
				} catch (Exception e) {
					e.printStackTrace();

					Row resultRow = resultsSheet.createRow(rowIndex++);
					resultRow.createCell(0).setCellValue(id);
					resultRow.createCell(1).setCellValue(city);
					resultRow.createCell(2).setCellValue(name);
					resultRow.createCell(3).setCellValue(size);
					resultRow.createCell(4).setCellValue(url);
					resultRow.createCell(5).setCellValue("NA");
					resultRow.createCell(6).setCellValue("NA");
					resultRow.createCell(7).setCellValue("NA");
					resultRow.createCell(8).setCellValue(uom);
					resultRow.createCell(9).setCellValue(" ");
					resultRow.createCell(10).setCellValue(" ");
					resultRow.createCell(11).setCellValue(" ");
					resultRow.createCell(12).setCellValue(" ");
					resultRow.createCell(13).setCellValue(" ");
					//resultRow.createCell(14).setCellValue("NA");

					System.out.println("Failed to extract data for URL: " + url);

				}
			}
			try {
				// for store the multiple we can use the time to store the multiple files
				SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
				String timestamp = dateFormat.format(new Date());
				String outputFilePath = ".\\Output\\Amazon__Ratings__OutputData_" + timestamp + ".xlsx";

				// Write results to Excel file
				FileOutputStream outFile = new FileOutputStream(outputFilePath);
				resultsWorkbook.write(outFile);
				outFile.close();

				System.out.println("Output file saved: " + outputFilePath);
			} catch (Exception e) {
				e.printStackTrace();
			}

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (driver != null) {
				System.out.println(" =======   DoNe DoNe Scraping DoNe  ========= ");
				driver.quit();
			}
		}
	}
}