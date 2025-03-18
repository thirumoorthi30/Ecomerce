package scheduler;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
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
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import CommonUtility.BlinkitId;

	
		import java.util.Calendar;
		import java.util.concurrent.Executors;
		import java.util.concurrent.ScheduledExecutorService;
		import java.util.concurrent.TimeUnit;

		public class scheduleDmart {
		    public static void main(String[] args) {
		    	ChromeOptions options = new ChromeOptions();
		    	options.addArguments("--headless"); // Run Chrome in headless mode
		    	options.addArguments("--disable-gpu"); // Disable GPU acceleration
		    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
		    	options.addArguments("--start-maximized");
		    	options.addArguments("--window-size=1920,1080");   //Set window size to full HD
		    	options.addArguments("--start-maximized");	

		        ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

		        // Schedule the task to run every day at 7:00 AM
		        Calendar now = Calendar.getInstance();
		        Calendar nextRunTime = Calendar.getInstance();
		        nextRunTime.set(Calendar.HOUR_OF_DAY, 5);
		        nextRunTime.set(Calendar.MINUTE, 0);
		        nextRunTime.set(Calendar.SECOND, 0);

		        long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
		        if (initialDelay < 0) {
		            initialDelay += 24 * 60 * 60 * 1000; // If it's already past 7 AM, schedule for the next day
		        }

		        scheduler.scheduleAtFixedRate(() -> {
		            try {
		                System.out.println("Starting web scraping task...");
		                scheduleDmart.runWebScraping();
		                System.out.println("Web scraping task completed.");
		            } catch (Exception e) {
		                e.printStackTrace();
		            }
		        }, initialDelay, 24 * 60 * 60 * 1000, TimeUnit.MILLISECONDS);
		    }


		    public static void runWebScraping() throws Exception{
		    
		        // Ensure this path points to the chromedriver.exe, not chrome.exe
		        // System.setProperty("webdriver.chrome.driver", "./Drivers/chromedriver.exe");

		         ChromeOptions options = new ChromeOptions();
//		         options.addArguments("--start-maximized");
//		         options.addArguments("--window-size=375,812");
//		         options.addArguments("--user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 13_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0 Mobile/15E148 Safari/604.1");

		         WebDriver driver = new ChromeDriver(options);
		         driver.manage().window().maximize();
		         WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

		         List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
		                 InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(),
		                 UOM = new ArrayList<>(), Mulitiplier = new ArrayList<>(),
		                 Pincode = new ArrayList<>(), NameForCheck = new ArrayList<>();

		         Workbook resultsWorkbook = new XSSFWorkbook();
		         Sheet resultsSheet = resultsWorkbook.createSheet("Results");
		         createHeaderRow(resultsSheet);
		 //This product is currently unavailable in your selected pincode
		         int rowIndex = 1;
		         String currentPin = null;

		         try (FileInputStream file = new FileInputStream(".\\input-data\\Dmart Input Data.xlsx");
		              Workbook urlsWorkbook = new XSSFWorkbook(file)) {

		             Sheet urlsSheet = urlsWorkbook.getSheet("DmartInput");
		             int rowCount = urlsSheet.getPhysicalNumberOfRows();

		             // Extract URLs from Excel
		             for (int i = 1; i < rowCount; i++) {
		                 Row row = urlsSheet.getRow(i);
		                 if (row.getCell(5) != null) {
		                     String url;
		                     if (row.getCell(5).getCellType() == CellType.STRING) {
		                         url = row.getCell(5).getStringCellValue();
		                     } else if (row.getCell(5).getCellType() == CellType.NUMERIC) {
		                         url = String.valueOf(row.getCell(5).getNumericCellValue());
		                     } else {
		                         url = "NA"; // Handle cases where the cell is neither numeric nor string
		                     }

		                     inputPid.add(row.getCell(0).getStringCellValue());
		                     InputCity.add(row.getCell(1).getStringCellValue());
		                     InputName.add(row.getCell(2).getStringCellValue());
		                     InputSize.add(row.getCell(3).getStringCellValue());
		                     NewProductCode.add(row.getCell(4).getStringCellValue());
		                     uRL.add(url);
		                     UOM.add(row.getCell(6).getStringCellValue());
		                     Mulitiplier.add(row.getCell(7).getStringCellValue());
		                     Pincode.add(row.getCell(9).getStringCellValue());
		                     NameForCheck.add(row.getCell(10).getStringCellValue());
		                 }
		             }

//		             Set<String> usedPincodes = new HashSet<>();

		             int ProductCOUNT = 0;

		             // Main data extraction logic
		             for (int i = 0; i < uRL.size(); i++) {
		                 String url = uRL.get(i);
		                 String locationSet = Pincode.get(i); // Get the current pincode from the list
		                 if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
		                     writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
		                             InputSize.get(i), "NA", url, "NA", "NA", "NA",
		                             "NA", Mulitiplier.get(i), "NA", "NA", "NA", "NA", "NA", "NA", NameForCheck.get(i));
		                     System.out.println("Skipped processing for URL: " + url);
		                     continue;
		                 }

		                 try {
		                     driver.get(url);
		                     Thread.sleep(2000);
		                     
		                   //  Thread.sleep(200000);
		                     if (currentPin == null || !currentPin.equals(locationSet)) {
		                     	Thread.sleep(2000);
		                     	driver.findElement(By.xpath("//div[@class='header_pincode__KryhE']")).click();
		                     	Thread.sleep(2000);
		                     	driver.findElement(By.xpath("//input[@id='pincodeInput']")).sendKeys(locationSet);
		                     	Thread.sleep(3000);
		                     	driver.findElement(By.xpath("(//div[@class='pincode-widget_pincode-right__TwcOu'])[1]")).click();
		                     	Thread.sleep(2000);
		                     	driver.findElement(By.xpath("//button[.='CONFIRM LOCATION']")).click();
		                     	Thread.sleep(2000);
		      		
		                     	currentPin = locationSet;
		                     	
		                     	driver.get(url);
		                          }  
		                     Thread.sleep(3000);


		                     System.out.println("PRODUCTCOUNT = " + ProductCOUNT);

		                     // Wait for the page to load
		                     String productId = extractProductId(url);
		                     System.out.println("Extracted Product ID: " + productId);
		                     ////div[@class='common_product-info__Y2P8l']//h1//span[1]
		                     
		                     String newName = " ";
		                     try {
		                     newName = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1"))).getText();
		                     System.out.println(newName);
		                     }
		                     catch (Exception e) {
		                     	newName = "NA";
		 					}
		                     // Extract weight
		                     String weight = extractWeight(driver);
		                     System.out.println("Extracted Weight: " + weight);
		                     // Extract MRP
		                     String mrpValue = extractMRP(driver);
		                     System.out.println(mrpValue);
		                     // Extract SP
		                     String spValue = extractSP(driver);
		                     System.out.println(spValue);
		                     // Extract offer
		                     String offerValue = extractOffer(driver);
		                     System.out.println(offerValue);

		                     int result;
		                     try {
		                         // Define the texts to check for
		                         String[] textsToCheck = {
		                                 "Out Of Stock",
		                                 "This product is currently unavailable in your selected pincode"
		                                 
		                         };

		                         // Get the page source
		                         String pageSource = driver.getPageSource();

		                         boolean isTextPresent = false;

		                         // Check for the presence of any of the texts
		                         for (String text : textsToCheck) {
		                             if (pageSource.contains(text)) {
		                                 isTextPresent = true;
		                                 break;
		                             }
		                         }
		                      


		                         // Determine the result based on the presence of the text
		                         result = isTextPresent ? 0 : 1;
		                         System.out.println(result);
		                     } catch (Exception e) {
		                         System.out.println("Error checking availability: " + e.getMessage());
		                         result = -1;
		                     }
//		                     BlinkitId screenshot = new BlinkitId();
//		                 	  try {
//		             				screenshot.screenshot(driver, "Dmart", productId);
//		             			} catch (Exception e) {
//		             				e.fillInStackTrace();
//		             			
//		             			}
		                     String availability = String.valueOf(result);

		                     ProductCOUNT++;

		                     // Write results to the results sheet
		                     writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
		                             InputSize.get(i), productId, url, newName, mrpValue, spValue,
		                             weight, String.valueOf(Mulitiplier.get(i)), availability, offerValue, "NA", "NA", "NA", "NA", NameForCheck.get(i));

		                     System.out.println("Data extracted for URL: " + url);
		                 } catch (Exception e) {
		                     e.printStackTrace();
		                     System.out.println("Failed to extract data for URL: " + url);
		                 }
		             }
		         } catch (Exception e) {
		             e.printStackTrace();
		             System.out.println("An error occurred during the extraction process.");
		         } finally {
		             // Write results to Excel file
		             saveResultsToExcel(resultsWorkbook);

		             // Clean up driver resources
		             if (driver != null) {
		                 System.out.println("Closing the driver.");
		                 driver.quit();
		             }
		         }
		     }

		     private static void createHeaderRow(Sheet sheet) {
		         Row headerRow = sheet.createRow(0);
		         String[] headers = {"InputPid", "InputCity", "InputName", "InputSize", "NewProductCode", "URL", "Name", "MRP",
		                 "SP", "UOM", "Multiplier", "Availability", "Offer", "Commands", "Remarks", "Correctness",
		                 "Percentage", "Name", "NameForCheck"};
		         for (int i = 0; i < headers.length; i++) {
		             headerRow.createCell(i).setCellValue(headers[i]);
		         }
		     }

		     private static void writeResults(Sheet sheet, int rowIndex, String id, String city, String name, String size,
		                                       String productId, String url, String newName, String mrpValue,
		                                       String spValue, String weight, String mulitiplier, String availability, String offerValue,
		                                       String command, String remark, String correctness, String percentage,
		                                       String nameCheck) {
		         Row resultRow = sheet.createRow(rowIndex);
		         resultRow.createCell(0).setCellValue(id);
		         resultRow.createCell(1).setCellValue(city);
		         resultRow.createCell(2).setCellValue(name);
		         resultRow.createCell(3).setCellValue(size);
		         resultRow.createCell(4).setCellValue(productId);
		         resultRow.createCell(5).setCellValue(url);
		         resultRow.createCell(6).setCellValue(newName);
		         resultRow.createCell(7).setCellValue(mrpValue);
		         resultRow.createCell(8).setCellValue(spValue);
		         resultRow.createCell(9).setCellValue(weight);
		         resultRow.createCell(10).setCellValue(mulitiplier);
		         resultRow.createCell(11).setCellValue(availability);
		         resultRow.createCell(12).setCellValue(offerValue);
		         resultRow.createCell(13).setCellValue(command);
		         resultRow.createCell(14).setCellValue(remark);
		         resultRow.createCell(15).setCellValue(correctness);
		         resultRow.createCell(16).setCellValue(percentage);
		         resultRow.createCell(17).setCellValue(nameCheck);
		     }

		     private static void saveResultsToExcel(Workbook resultsWorkbook) {
		         try {
		             SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
		             String timestamp = dateFormat.format(new Date());
		             String outputFilePath = ".\\Output\\DMART_OutputData_" + timestamp + ".xlsx";

		             // Write results to Excel file
		             try (FileOutputStream outFile = new FileOutputStream(outputFilePath)) {
		                 resultsWorkbook.write(outFile);
		             }

		             System.out.println("Output file saved: " + outputFilePath);
		         } catch (Exception e) {
		             e.printStackTrace();
		             System.out.println("Failed to save the output file.");
		         }
		     }
		     
		     

		     private static String extractMRP(WebDriver driver) {
		     	try {
		             List<WebElement> elements = driver.findElements(By.xpath("//span[contains(., 'MRP')]/span"));
		             if (elements.isEmpty()) {
		                 return "NA";  // MRP value not found
		             }
		             String text = elements.get(0).getText();
		             return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
		         } catch (Exception e) {
		             return "NA";  // In case of any other exception
		         }

		     }

		     private static String extractSP(WebDriver driver) {
		     	try {
		             List<WebElement> elements = driver.findElements(By.xpath("//span[contains(., 'DMart')]/span"));
		             if (elements.isEmpty()) {
		                 return "NA";  // SP value not found
		             }
		             String text = elements.get(0).getText();
		             return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
		         } catch (Exception e) {
		             return "NA";  // In case of any other exception
		         }
		         
		     }

		     private static String extractOffer(WebDriver driver) {
		     	try {
		             List<WebElement> elements = driver.findElements(By.xpath("//div[@class='price-details-component_saveHighlighter__FIIS_']//div"));
		             if (elements.isEmpty()) {
		                 return "NA";  // Offer value not found
		             }
		             String text = elements.get(0).getText();
		             return isValidValue(text) ? text.trim() : "NA";
		         } catch (Exception e) {
		             return "NA";  // In case of any other exception
		         }
		      
		     }
		     private static String extractProductId(String url) {
		         // Find the index of "selectedProd="
		         int startIndex = url.indexOf("selectedProd=") + 13; // Add 13 to move past "selectedProd="
		         
		         // Find the index of the next "&" or end of string
		         int endIndex = url.indexOf("&", startIndex);
		         
		         // If no "&" is found, extract till the end of the string
		         if (endIndex == -1) {
		             endIndex = url.length();
		         }
		         
		         // Extract and return the product ID using substring
		         return url.substring(startIndex, endIndex);
		     }

		     
		     private static String extractWeight(WebDriver driver) {
		     	try {
		             List<WebElement> elements = driver.findElements(By.xpath("//h1//span[2]"));
		             if (elements.isEmpty()) {
		                 return "NA";  // Weight value not found
		             }
		             String text = elements.get(0).getText();
		             return isValidValue(text) ? text.replace(":", "").trim() : "NA";
		         } catch (Exception e) {
		             return "NA";  // In case of any other exception
		         }

		     }
		     private static boolean isValidValue(String value) {
		         return value != null && !value.isEmpty() && !value.equals("₹");
		     }
		    

}

