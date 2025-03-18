package Top3Result;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class amazonProductScraper {

    public static void main(String[] args) {
        // Generate timestamp
        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String timestamp = currentTime.format(formatter);

        String inputFile = ".\\input-data\\Product Search.xlsx";
        String outputFile = ".\\Output\\Amazon_Product_Search_OutputData_  " + timestamp + " .xlsx";

        // Set Chrome driver path
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");

        // Initialize Chrome options
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");

        // Initialize WebDriver
        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        try {
            // Initialize output Excel workbook
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Results");
            
            
            Row headerRow = outputSheet.createRow(0);
            headerRow.createCell(0).setCellValue("Input Product ID");
            headerRow.createCell(1).setCellValue("Input Product Name");
            headerRow.createCell(2).setCellValue("Input Product Uom");
            headerRow.createCell(3).setCellValue("Product URL");
            headerRow.createCell(4).setCellValue("Product Name");
            headerRow.createCell(5).setCellValue("MRP");
            headerRow.createCell(6).setCellValue("SP");
            

            // Read input Excel sheet
            Workbook workbook = WorkbookFactory.create(new File(inputFile));
            Sheet sheet = workbook.getSheetAt(0); // Assuming input is in first sheet
            
            //String originalSp= "";

            for (Row row : sheet) {
            	
            	if (row == null || row.getCell(0) == null || row.getCell(0).getStringCellValue().isEmpty()) {
                    break;
                }
                String iPid = row.getCell(0).getStringCellValue();
                String productName = row.getCell(1).getStringCellValue();
                String uom = row.getCell(2).getStringCellValue();
                
                // Assuming product name is in first column

                driver.get("https://www.amazon.in/?&tag=googhydrabk1-21&ref=pd_sl_7hz2t19t5c_e&adgrpid=155259815513&hvpone=&hvptwo=&hvadid=674842289437&hvpos=&hvnetw=g&hvrand=12817599141667678899&hvqmt=e&hvdev=c&hvdvcmdl=&hvlocint=&hvlocphy=9148906&hvtargid=kwd-10573980&hydadcr=14453_2316415&gad_source=1");
                
                //PINCODE
               int m = 0;
                
            	if(m == 0 || m == 1) {
    			WebElement location = driver.findElement(By.id("nav-global-location-popover-link"));
    			location.click();
    			String tempPinNumber = "";
    			for (int j = 0; j < 150; j++) {
    				try {
    					driver.findElement(	
    							By.xpath("//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']"))
    							.sendKeys(Keys.ENTER);
    					
    					driver.findElement(
    							By.xpath("//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']")).clear();
    					
    					//System.out.println("print the crt pin number" + InputPin);
    					
    					//String crtPin = InputPin;
    					driver.findElement(
    							By.xpath("//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']"))
    							.sendKeys("110001");
    					
    				/*	driver.findElement(
    							By.xpath("//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']"))
    							.sendKeys(Keys.ENTER);   */
    					
    					for (int k = 0; k <= 50; k++) {
    						try {
    							tempPinNumber = driver.findElement(By.xpath(
    									"//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']"))
    									.getAttribute("value");
    							if (tempPinNumber.equals("110001")) {
    								break;
    							}
    						} catch (Exception e) {
    							if (m == 50) {
    								Assert.fail(e.getMessage());
    							}
    						}
    					}
    					driver.findElement(By.xpath("//span[contains(text(),'Apply')]")).click();
    					
    					
    					
    				
    					
    					break;
    				} catch (Exception e) {
    					e.getCause();
    					if (j == 300) {	
    						Assert.fail(e.getMessage());
    					}
    				}
    			}  
    		}   
                
                
                
                
                
            	 WebElement searchInput = driver.findElement(By.id("twotabsearchtextbox"));
                 searchInput.sendKeys(productName);
                 searchInput.sendKeys(Keys.RETURN);

                 // Wait for the search results to load
                 Thread.sleep(4000);

                 // Find all product elements in the search results
                 List<WebElement> productElements = driver.findElements(By.xpath("//div[@data-component-type='s-search-result']"));

                 int count = 0;
                 
                 String originalSp = " ";
                 String originalMrp = "";
                 String productNameText = "";
                 
                 String productUrl = "";
                 
                 // Loop through the product elements and select non-sponsored products
                 for (WebElement productElement : productElements) {
                     try {
                         // Check if the product is sponsored
                         productElement.findElement(By.className("puis-label-popover-default"));
                         continue; // Skip sponsored products
                     } catch (org.openqa.selenium.NoSuchElementException ignored) {
                         // Product is not sponsored
                         count++;

                         // Get necessary data for the product
                         WebElement productNameElement = productElement.findElement(By.xpath(".//h2/a/span"));
                         productNameText = productNameElement.getText();

                         WebElement productLink = productElement.findElement(By.xpath(".//h2/a"));
                         productUrl = productLink.getAttribute("href");
                         
                         
                        
                         	try {
                         WebElement mrp = productElement.findElement(By.xpath(".//div[@class='a-section aok-inline-block']//span[@class='a-price a-text-price']//span[@aria-hidden]"));
                          originalMrp = mrp.getText().replaceAll("[^\\d.]+", "");
                         	}
                         	catch(Exception e) {
                         		try {
                         			WebElement mrp = productElement.findElement(By.xpath(".//span[@class='a-price-whole']"));
                                    originalMrp = mrp.getText().replaceAll("[^\\d.]+", "");
                         			
                         		}
                         		catch (Exception a) {
                         			try {
                         				WebElement mrp1 = productElement.findElement(By.xpath("//*[@id=\"search\"]/div[1]/div[1]/div/span[1]/div[1]/div[5]/div/div/span/div/div/div[2]/div[3]/div/div[1]/a/span[1]/span[2]/span[2]"));
                                        originalMrp = mrp1.getText().replaceAll("[^\\d.]+", "");
                         			}
                         			catch(NoSuchElementException y) {
                         				originalMrp = "NA";
                         			  }
                         		}
                         		
                         	}
                         	
                         	try {
                         WebElement sp = productElement.findElement(By.xpath(".//span[@class='a-price-whole']"));
                         originalSp = sp.getText().replace("₹", "");
                         
                         	}
                         	catch(NoSuchElementException n) {
                         		originalSp = "NA";
                         	   }
                         	}

                         // Write data to output Excel sheet
                         Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                         
                         outputRow.createCell(0).setCellValue(iPid);
                         outputRow.createCell(1).setCellValue(productName);
                         outputRow.createCell(2).setCellValue(uom);
                         outputRow.createCell(3).setCellValue(productUrl);
                         outputRow.createCell(4).setCellValue(productNameText);
                         outputRow.createCell(5).setCellValue(originalMrp);
                         outputRow.createCell(6).setCellValue(originalSp);

                         System.out.println("Product Name: " + productNameText);
                         System.out.println("Product URL: " + productUrl);
                         System.out.println("MRP: " + originalMrp);
                         System.out.println("SP: " + originalSp);

                         if (count >= 3) {
                             break; // Found three non-sponsored products, exit loop
                         }
                     }
                    
                 }
             


            // Save output Excel file
			
			  FileOutputStream fileOut = new FileOutputStream(outputFile);
			  outputWorkbook.write(fileOut); 
			  fileOut.close();
			 
			  System.out.println("DoNE Scrapping DoNE");
			  

            // Close the WebDriver
            driver.quit();

        } 
    catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
}