package Shopping;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.testng.Assert;

import CommonUtility.BlinkitId;


import java.util.Calendar;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

public class T1mg {
	public static void main(String[] args) {
		
		//      System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		ChromeOptions options = new ChromeOptions();
		
		int count = 0;

		String spValue = "";
		String finalSp = "";
		String offerValue = "NA";
		String newName = null;
		String mrpValue = null;
		String originalMrp1 = " ";
		String originalMrp2 = " ";
		String originalMrp3 = " ";
		String originalSp1 = " ";
		String originalSp2 = " ";
		String NewAvailability1 = " ";
		String originalName1=" ";
		String marketName1=" ";
		String saltName1=" ";
		String storeName1=" ";
		
		
		try {
			// Read URLs from Excel file
			String filePath = ".\\input-data\\T1mg.xlsx";
			FileInputStream file = new FileInputStream(filePath);
			Workbook urlsWorkbook = new XSSFWorkbook(file);
			Sheet urlsSheet = urlsWorkbook.getSheet("Sheet2");
			int rowCount = urlsSheet.getPhysicalNumberOfRows();

			List<String> InputName = new ArrayList<>(),uRL = new ArrayList<>();

			// Extract URLs from Excel
			for (int i = 0; i < rowCount; i++) {
				Row row = urlsSheet.getRow(i);


				if (i == 0) {
					continue;
				}    

				Cell inputNameCell = row.getCell(0);
				Cell urlCell = row.getCell(1);
				Cell pinCodeCell = row.getCell(2);        

				//   Cell urlCell = row.getCell(0);
				//  Cell urlCell = row.getCell(0);
				// Cell idCell = row.getCell(1);
				// Cell offerCell = row.getCell(2);

				if (urlCell != null && urlCell.getCellType() == CellType.STRING) {
					String url = urlCell.getStringCellValue();
					String name = (inputNameCell != null && inputNameCell.getCellType() == CellType.STRING) ? inputNameCell.getStringCellValue() : "";
					String locationSet = (pinCodeCell != null && pinCodeCell.getCellType() == CellType.STRING) ? pinCodeCell.getStringCellValue() : "";

					
					InputName.add(name);
					uRL.add(url);
					

					/*
					 * uRL.add(url); ids.add(id); offers.add(offer);
					 */

				}
			}
			// Create Excel workbook for storing results
			Workbook resultsWorkbook = new XSSFWorkbook();
			Sheet resultsSheet = resultsWorkbook.createSheet("Results");

			Row headerRow = resultsSheet.createRow(0);


			
			headerRow.createCell(0).setCellValue("InputName");
			headerRow.createCell(1).setCellValue("URL");
			headerRow.createCell(2).setCellValue("Name");
			headerRow.createCell(3).setCellValue("Marketer");
			headerRow.createCell(4).setCellValue("Salt");
			headerRow.createCell(5).setCellValue("Storage");
			headerRow.createCell(6).setCellValue("MRP");
			headerRow.createCell(7).setCellValue("sp");
			headerRow.createCell(8).setCellValue("Offer");
			
			int rowIndex = 1;
			int headercount = 0;
			String currentPin = null;

			for (int i = 0; i < uRL.size(); i++) {
				
				String name = InputName.get(i);
				String url = uRL.get(i);
				


				try {

					if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
						// Set "NA" values in all three columns
						Row resultRow = resultsSheet.createRow(rowIndex++);
					
						resultRow.createCell(0).setCellValue(name);
						resultRow.createCell(1).setCellValue(url);
						resultRow.createCell(2).setCellValue("NA");
						resultRow.createCell(3).setCellValue("NA");
						resultRow.createCell(4).setCellValue("NA");
						resultRow.createCell(5).setCellValue("NA");
						resultRow.createCell(6).setCellValue("NA");
						resultRow.createCell(7).setCellValue("NA");
						resultRow.createCell(8).setCellValue("NA");	
						
						System.out.println("Skipped processing for URL: " + url);
						continue; // Skip to the next iteration
					}

					driver.get(url);
					driver.manage().window().maximize();

					//Name
					try {
						WebElement ProName = driver.findElement(By.xpath("//h1"));
						originalName1 = ProName.getText();
						System.out.println(originalName1);

					}catch (Exception d) {
						// TODO: handle exception
					}

					System.out.println("headercount = " + headercount);

					headercount++;
					
					//Marketer
					try {
						WebElement Marketer = driver.findElement(By.xpath("(//div[@class='DrugHeader__meta___B3BcU']//div//a)[1]"));
						marketName1 = Marketer.getText();
						System.out.println(marketName1);

					}catch (Exception d) {
						// TODO: handle exception
					}
					
					//Salt
					try {
						WebElement Salter = driver.findElement(By.xpath("(//div[@class='DrugHeader__meta___B3BcU']//div//a)[2]"));
						saltName1 = Salter.getText();
						System.out.println(saltName1);

					}catch (Exception v) {
						// TODO: handle exception
					}
					
					//storage
					try {
						WebElement store = driver.findElement(By.xpath("(//div[@class='DrugHeader__meta___B3BcU']//div/following-sibling::div)[3]"));
						storeName1 = store.getText();
						System.out.println(storeName1);

					}catch (Exception l) {
						// TODO: handle exception
					}
					
					
					Thread.sleep(5000);
					
					try {
						WebElement sp = driver.findElement(By.xpath("//div[@class='DrugPriceBox__price___dj2lv']"));
						originalSp1 = sp.getText();
						spValue =  originalSp1.replace("₹", "").replace("Inclusive of all taxes", "").replace("MRP", "");
						System.out.println(spValue);

					}
					catch(Exception e) {
						try {
							WebElement sp = driver.findElement(By.xpath("(//span[@class='PriceBoxPlanOption__offer-price___3v9x8 PriceBoxPlanOption__offer-price-cp___2QPU_'])[1]"));
							originalSp1 = sp.getText();
							spValue =  originalSp1.replace("₹", "").replace("Inclusive of all taxes", "").replace("MRP", "");
							System.out.println(spValue);
						}
						catch (Exception e2) {
							spValue = "NA";
						}}
					
					            
					//MRP
					Thread.sleep(2000);
					try {
						WebElement mrp = driver.findElement(By.xpath("//span[@class='DrugPriceBox__slashed-price___2UGqd']"));
						originalMrp1 = mrp.getText();
						mrpValue = originalMrp1.replace("₹", "");
						System.out.println(mrpValue);

					} 	catch (Exception e) {
						try{
							WebElement mrp = driver.findElement(By.xpath("//span[@class='PriceBoxPlanOption__margin-right-4___2aqFt PriceBoxPlanOption__stike___pDQVN']"));
							originalMrp1 = mrp.getText();
							mrpValue = originalMrp1.replace("₹", "");
							System.out.println(mrpValue);
						}catch (NoSuchElementException q) {
							mrpValue = spValue;
						}
					}

					

				
					// OFFER

					if(mrpValue.equals(spValue)){
						offerValue = "NA";
					}
					else {
						try {
							WebElement offer = driver.findElement(By.xpath("//span[@class='DrugPriceBox__slashed-percent___G92cz']"));
							String offer1 = offer.getText();
							String offer2 = offer1.replace("% OFF", "% Off");

							offerValue = offer2;
							System.out.println(offerValue);

						}
						catch(Exception e) {
							WebElement offer = driver.findElement(By.xpath("//span[@class='PriceBoxPlanOption__discount___iN_jm']"));
							String offer1 = offer.getText();
							String offer2 = offer1.replace("% OFF", "% Off").replace("GET", "");

							offerValue = offer2;
							System.out.println(offerValue);

						}


					}

				

					Row resultRow = resultsSheet.createRow(rowIndex++);

					
					resultRow.createCell(0).setCellValue(name); 
					resultRow.createCell(1).setCellValue(url);
					resultRow.createCell(2).setCellValue(originalName1);
					resultRow.createCell(3).setCellValue(marketName1);
					resultRow.createCell(4).setCellValue(saltName1);
					resultRow.createCell(5).setCellValue(storeName1);
					resultRow.createCell(6).setCellValue(mrpValue);
					resultRow.createCell(7).setCellValue(spValue);
					resultRow.createCell(8).setCellValue(offerValue);
					


					System.out.println("Data extracted for URL: " + url);
				} catch (Exception e) {
					e.printStackTrace();

					Row resultRow = resultsSheet.createRow(rowIndex++);
					
					resultRow.createCell(0).setCellValue(name);
					resultRow.createCell(1).setCellValue(url);
					resultRow.createCell(2).setCellValue("NA");
					resultRow.createCell(3).setCellValue("NA");
					resultRow.createCell(4).setCellValue("NA");
					resultRow.createCell(5).setCellValue("NA");
					resultRow.createCell(6).setCellValue("NA");
					resultRow.createCell(7).setCellValue("NA");
					resultRow.createCell(8).setCellValue("NA");

					System.out.println("Failed to extract data for URL: " + url);

				}
			}
			try {
				// for store the multiple we can use the time to store the multiple files
				SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
				String timestamp = dateFormat.format(new Date());
				String outputFilePath = ".\\Output\\T1mg_OutputData_" + timestamp + ".xlsx";


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
				System.out.println("DoNe DoNe Scraping DoNe");
				driver.quit();
			}
		}
	}
}    
