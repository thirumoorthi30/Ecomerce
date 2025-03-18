package BBFnv;


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

		public class scheduleBB2 {
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
		        nextRunTime.set(Calendar.HOUR_OF_DAY, 2);
		        nextRunTime.set(Calendar.MINUTE, 10);
		        nextRunTime.set(Calendar.SECOND, 0);

		        long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
		        if (initialDelay < 0) {
		            initialDelay += 24 * 60 * 60 * 1000; // If it's already past 7 AM, schedule for the next day
		        }

		        scheduler.scheduleAtFixedRate(() -> {
		            try {
		                System.out.println("Starting web scraping task...");
		                scheduleBB2.runWebScraping();
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
//		             options.addArguments("--start-maximized");
//		             options.addArguments("--window-size=375,812");
//		             options.addArguments("--user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 13_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0 Mobile/15E148 Safari/604.1");

		    		WebDriver driver = new ChromeDriver(options);
		    		driver.manage().window().maximize();
		    		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

		    		List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
		    				InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(),
		    				UOM = new ArrayList<>(), Mulitiplier = new ArrayList<>(), Pincode = new ArrayList<>(),
		    				NameForCheck = new ArrayList<>();

		    		Workbook resultsWorkbook = new XSSFWorkbook();
		    		Sheet resultsSheet = resultsWorkbook.createSheet("Results");
		    		createHeaderRow(resultsSheet);
		    		// This product is currently unavailable in your selected pincode
		    		int rowIndex = 1;
		    		String currentPin = null;

		    		try (FileInputStream file = new FileInputStream(".\\input-data\\BB28 Input Data New.xlsx");
		    				Workbook urlsWorkbook = new XSSFWorkbook(file)) {

		    			Sheet urlsSheet = urlsWorkbook.getSheet("BBNew2");
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

//		                 Set<String> usedPincodes = new HashSet<>();

		    			int ProductCOUNT = 0;

		    			// Main data extraction logic
		    			for (int i = 0; i < uRL.size(); i++) {
		    				String url = uRL.get(i);
		    				String locationSet = Pincode.get(i); // Get the current pincode from the list
		    				if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
		    					writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
		    							InputSize.get(i), NewProductCode.get(i), url, "NA", "NA", "NA", UOM.get(i),
		    							Mulitiplier.get(i), "NA", "NA", "NA", "NA", "NA", "NA", NameForCheck.get(i));
		    					System.out.println("Skipped processing for URL: " + url);
		    					continue;
		    				}

		    				try {
		    					driver.get(url);
		    					Thread.sleep(2000);

		    					// Thread.sleep(200000);
		    					WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(10));

		    					if (currentPin == null || !currentPin.equals(locationSet)) {
		    						// Wait for and click the button to open the location setting
		    						try {
		    							// Navigate to BigBasket website
		    							driver.get("https://www.bigbasket.com/");

		    							// Wait for the location input field to be clickable (based on your HTML
		    							// snippet)
		    							WebDriverWait wait11 = new WebDriverWait(driver, Duration.ofSeconds(10));

		    							try {

		    							WebElement button = wait11.until(ExpectedConditions
		    									.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[2]")));
		    							button.click();
		    							
		    							}catch (Exception e) {
		    								WebElement button = wait11.until(ExpectedConditions
			    									.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[1]")));
			    							button.click();
										}
		    							
		    							try {
		    								WebElement locationInput = wait.until(ExpectedConditions.elementToBeClickable(By
		    										.xpath("//div[@class='flex flex-col absolute right-0 top-full mt-1.5 bg-white rounded-2xs outline-none z-max w-74 xl:w-90 scale-100']//div[@class='AddressDropdown___StyledDiv-sc-i4k67t-7 eXGbTp']//input")));

		    								// Click on the location input field
		    								locationInput.click();
		    								locationInput.sendKeys(locationSet);
		    							} catch (Exception e) {

		    								List<WebElement> searchInputElement = driver.findElements(
		    										By.xpath("//input[@placeholder='Search for area or street name']"));
		    								System.out.println("Search box size " + searchInputElement.size());
		    								for (WebElement searchInput : searchInputElement) {
		    									try {
		    										searchInput.click();
		    										searchInput.sendKeys(locationSet);
		    									} catch (Exception e1) {
		    										e1.printStackTrace();
		    									}

		    								}
//		                                	 
		    							}

		    							// Wait for suggestions to appear
		    							// Click the first XPath
		    						
		    							WebElement firstElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ul[@class='overscroll-contain p-2.5']//li[1]")));
		    							firstElement.click();

		    							Thread.sleep(2000);
		    							
		    							// Wait for the message to appear
		    							WebElement message = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[contains(text(),'The selected city is not serviceable at the moment')]")));

		    							if (message.isDisplayed()) {
		    								
		    								driver.get("https://www.bigbasket.com/");


			    							try {

			    							WebElement button = wait1.until(ExpectedConditions
			    									.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[2]")));
			    							button.click();
			    							
			    							}catch (Exception e) {
			    								WebElement button = wait1.until(ExpectedConditions
				    									.elementToBeClickable(By.xpath("(//div[@class='flex w-full']//button)[1]")));
				    							button.click();
											}
			    							
			    							try {
			    								WebElement locationInput = wait1.until(ExpectedConditions.elementToBeClickable(By
			    										.xpath("//div[@class='flex flex-col absolute right-0 top-full mt-1.5 bg-white rounded-2xs outline-none z-max w-74 xl:w-90 scale-100']//div[@class='AddressDropdown___StyledDiv-sc-i4k67t-7 eXGbTp']//input")));

			    								// Click on the location input field
			    								locationInput.click();
			    								locationInput.sendKeys(locationSet);
			    							} catch (Exception e) {

			    								List<WebElement> searchInputElement = driver.findElements(
			    										By.xpath("//input[@placeholder='Search for area or street name']"));
			    								System.out.println("Search box size " + searchInputElement.size());
			    								for (WebElement searchInput : searchInputElement) {
			    									try {
			    										searchInput.click();
			    										searchInput.sendKeys(locationSet);
			    									} catch (Exception e1) {
			    										e1.printStackTrace();
			    									}

			    								}
			    
			    							}
			    						
		    							    // If the message appears, click the second XPath
		    							    WebElement secondElement = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//ul[@class='overscroll-contain p-2.5']//li[2]")));
		    							    secondElement.click();
		    							}

		    						
		    						
		    							System.out.println("Location set to: " + locationSet);
		    							
		    							currentPin = locationSet;

		    						} catch (Exception e) {
		    							e.printStackTrace();
		    							currentPin = locationSet;

		    							driver.get(url);
		    						}
		    					}

		    					Thread.sleep(3000);

		    					System.out.println("PRODUCTCOUNT = " + ProductCOUNT);
		    					driver.get(url);

		    					// Wait for the page to load
		    					// String productId = extractProductId(url);
		    					// System.out.println("Extracted Product ID: " + productId);
		    					//// div[@class='common_product-info__Y2P8l']//h1//span[1]

		    					String newName = " ";
		    					try {
		    						newName = wait.until(ExpectedConditions.visibilityOfElementLocated(
		    								By.xpath("//h1[@class='Description___StyledH-sc-82a36a-2 bofYPK']"))).getText();
		    						System.out.println(newName);
		    					} catch (Exception e) {
		    						newName = "NA";
		    					}
		    					// Extract weight
		    					// String weight = extractWeight(driver);
		    					// System.out.println("Extracted Weight: " + weight);
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
		    						String[] textsToCheck = { "Add to basket"

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
		    						result = isTextPresent ? 1 : 0;
		    						System.out.println(result);
		    					} catch (Exception e) {
		    						System.out.println("Error checking availability: " + e.getMessage());
		    						result = -1;
		    					}
//		                         BlinkitId screenshot = new BlinkitId();
//		                     	  try {
//		                 				screenshot.screenshot(driver, "Dmart", productId);
//		                 			} catch (Exception e) {
//		                 				e.fillInStackTrace();
//		                 			
//		                 			}
		    					String availability = String.valueOf(result);

		    					ProductCOUNT++;

		    					// Write results to the results sheet
		    					writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
		    							InputSize.get(i), NewProductCode.get(i), url, newName, mrpValue, spValue, UOM.get(i),
		    							String.valueOf(Mulitiplier.get(i)), availability, offerValue, " ", " ", " ", " ",
		    							NameForCheck.get(i));

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
		    		String[] headers = { "InputPid", "InputCity", "InputName", "InputSize", "NewProductCode", "URL", "Name", "MRP",
		    				"SP", "UOM", "Multiplier", "Availability", "Offer", "Commands", "Remarks", "Correctness", "Percentage",
		    				"Name", "NameForCheck" };
		    		for (int i = 0; i < headers.length; i++) {
		    			headerRow.createCell(i).setCellValue(headers[i]);
		    		}
		    	}

		    	private static void writeResults(Sheet sheet, int rowIndex, String id, String city, String name, String size,
		    			String productId, String url, String newName, String mrpValue, String spValue, String weight,
		    			String mulitiplier, String availability, String offerValue, String command, String remark,
		    			String correctness, String percentage, String nameCheck) {
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
		    			String outputFilePath = ".\\Output\\BB_OutputData_SecondHalf" + timestamp + ".xlsx";

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
		    			List<WebElement> elements = driver
		    					.findElements(By.xpath("//td[contains(text(), 'MRP:')]/following-sibling::td"));
		    			if (elements.isEmpty()) {
		    				return "NA"; // MRP value not found
		    			}
		    			String text = elements.get(0).getText();
		    			return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
		    		} catch (Exception e) {
		    			return "NA"; // In case of any other exception
		    		}

		    	}

		    	private static String extractSP(WebDriver driver) {
		    		try {
		    			List<WebElement> elements = driver.findElements(By.xpath("//td[contains(text(), 'Price:')]"));
		    			if (elements.isEmpty()) {
		    				return "NA"; // SP value not found
		    			}
		    			String text = elements.get(0).getText();
		    			return isValidValue(text) ? text.replace("Price: ₹", "").trim() : "NA";
		    		} catch (Exception e) {
		    			return "NA"; // In case of any other exception
		    		}

		    	}

		    	private static String extractOffer(WebDriver driver) {
		    		try {
		    			List<WebElement> elements = driver
		    					.findElements(By.xpath("//td[contains(text(), 'You Save:')]/following-sibling::td"));
		    			if (elements.isEmpty()) {
		    				return "NA"; // Offer value not found
		    			}
		    			String text = elements.get(0).getText();
		    			return isValidValue(text) ? text.replace("OFF", "Off").trim() : "NA";
		    		} catch (Exception e) {
		    			return "NA"; // In case of any other exception
		    		}

		    	}

		    	/*
		    	 * private static String extractProductId(String url) { // Find the index of
		    	 * "selectedProd=" int startIndex = url.indexOf("selectedProd=") + 13; // Add 13
		    	 * to move past "selectedProd="
		    	 * 
		    	 * // Find the index of the next "&" or end of string int endIndex =
		    	 * url.indexOf("&", startIndex);
		    	 * 
		    	 * // If no "&" is found, extract till the end of the string if (endIndex == -1)
		    	 * { endIndex = url.length(); }
		    	 * 
		    	 * // Extract and return the product ID using substring return
		    	 * url.substring(startIndex, endIndex); }
		    	 * 
		    	 * 
		    	 * private static String extractWeight(WebDriver driver) { try {
		    	 * List<WebElement> elements = driver.findElements(By.xpath("//h1//span[2]"));
		    	 * if (elements.isEmpty()) { return "NA"; // Weight value not found } String
		    	 * text = elements.get(0).getText(); return isValidValue(text) ?
		    	 * text.replace(":", "").trim() : "NA"; } catch (Exception e) { return "NA"; //
		    	 * In case of any other exception }
		    	 * 
		    	 * }
		    	 */
		    	private static boolean isValidValue(String value) {
		    		return value != null && !value.isEmpty() && !value.equals("₹");
		    	}

		

		        }
		    