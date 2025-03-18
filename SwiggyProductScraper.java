package Shopping;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;


// This is the full code for SWIGGY Top 3 search But it is not working Stale exception error issue... so will have to use the other code SwiggyproductlinkScrapper..
//then use the normal demo m code to scrape the details....


public class SwiggyProductScraper {

    public static void main(String[] args) {
        // Generate timestamp
        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String timestamp = currentTime.format(formatter);

        String inputFile = ".\\input-data\\Product Search.xlsx";
        String outputFile = ".\\Output\\Swiggy_Pro_Ser_OutputData " + timestamp + " .xlsx";

        // Set Chrome driver path
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");

        // Initialize Chrome options
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");

        // Initialize WebDriver
        WebDriver driver = new ChromeDriver(options);
       // WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        try {
            // Initialize output Excel workbook
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Results");
            
            Row headerRow = outputSheet.createRow(0);
            headerRow.createCell(0).setCellValue("PID");
            headerRow.createCell(1).setCellValue("CITY");
            headerRow.createCell(2).setCellValue("Input Product Name");
            headerRow.createCell(3).setCellValue("UOM");
            headerRow.createCell(4).setCellValue("Product URL");
            headerRow.createCell(5).setCellValue("Product Name");
            headerRow.createCell(6).setCellValue("MRP");
            headerRow.createCell(7).setCellValue("SP");
            headerRow.createCell(8).setCellValue("UOM");
            headerRow.createCell(9).setCellValue("Multiplier");
            headerRow.createCell(10).setCellValue("NAME");

            // Read input Excel sheet
            Workbook workbook = WorkbookFactory.create(new File(inputFile));
            Sheet sheet = workbook.getSheetAt(0); // Assuming input is in first sheet

            for (Row row : sheet) {
                if (row == null || row.getCell(0) == null || row.getCell(0).getStringCellValue().isEmpty()) {
                    break;
                }
                
                String pId = row.getCell(0).getStringCellValue();
                String city = row.getCell(1).getStringCellValue();
                String productName = row.getCell(2).getStringCellValue();
                String uom = row.getCell(3).getStringCellValue();
                
                
                System.out.println(pId);
                
                
                if(pId.equals("407473")) {
               
                driver.get("https://www.swiggy.com/instamart");
                
               // driver.manage().deleteAllCookies();
                
                Thread.sleep(20000);

                
           
                try {
                for(int i=0; i<70; i++) {
                	 WebElement searchInput = driver.findElement(By.xpath(" .//div[@class='_1bOXc']"));
                	 searchInput.click();
                	 
                //searchInput.sendKeys(Keys.ENTER);
                	 break;
                }
                }
                catch(StaleElementReferenceException t) {
                	
                }
                Thread.sleep(5000);
                
                // Find the search bar and search for each product
               WebElement searchInputButton = driver.findElement(By.xpath(" .//button[@class='icon-search _1WtbA _2rOBm']"));
               searchInputButton.click();
         
                }
                
                Thread.sleep(4000);
             
               
             //   searchInput.click();
                
                WebElement searchInputClick = driver.findElement(By.xpath(" .//input[@class='cK7Br']"));
                
                searchInputClick.click();
                
                Thread.sleep(2000);
                
                searchInputClick.clear(); 	
                
                Thread.sleep(2000);
                
                searchInputClick.sendKeys(productName);
                Thread.sleep(3000);// Enter the product name
                searchInputClick.sendKeys(Keys.ENTER); // Submit the search query
                
           
                Thread.sleep(7000);

                // Find all product elements in the search results
                List<WebElement> productElements = driver.findElements(By.xpath(" .//div[@class='XjYJe']"));
                
                System.out.println(productElements);

            
                String productUrl = " ";
                String name1 = " ";
                String name2 = " ";
                String fullProductName = " ";
                String originalMrp = " ";
                String spValue = " ";
                String originalUom = " ";
                
                // Flag to check if the third product was found
                boolean thirdProductFound = false;
                
                int count = Math.min(productElements.size(), 3);
                
                System.out.println(count);
                
                // Loop through the first three product elements
                for (int i = 0; i < count; i++) {
                    // Get the ith product element
                	 WebElement productElement = null;
                	    try {
                	        productElement = productElements.get(i);
                	    } catch (IndexOutOfBoundsException e) {
                	        
                	        System.err.println("Index out of bounds: " + e.getMessage());
                	        continue; 
                	    }
          
                    Thread.sleep(2000);
                    
                    List<WebElement> productElementsList = driver.findElements(By.xpath(" .//div[@class='_1OklP _1TqJT']"));
                    
                    List<WebElement> productElements1 = driver.findElements(By.xpath(" .//div[@class='XjYJe']"));
                 // Loop through each product element in the list
                    for (WebElement productElementInside : productElements1) {
                    	
                    	
                    	try {
                            // Click on the product element
                            productElementInside.click();
                            
                            
                            Thread.sleep(2000);
     
                            
                            // Add logging statements to confirm successful click and retrieve URL
                            System.out.println("productElementInside clicked successfully.");
                        
                        } catch (StaleElementReferenceException e) {
                            // Handle the stale element exception by refreshing the list of product elements
                            productElements = driver.findElements(By.xpath(".//div[@class='XjYJe']"));
                        } catch (Exception e) {
                            // Handle any other exceptions that occur
                            e.printStackTrace();
                            // Add logging statement to indicate failure
                            System.out.println("Failed to click product element.");
                        }
                    	
                    	Thread.sleep(3000);
                        
                        try {
                        	productUrl = driver.getCurrentUrl();
                        }
                        catch(Exception j) {
                        	j.printStackTrace();
                        	System.out.println("Product Generated successfully.");
                        	System.out.println("Product URL: " + productUrl);
                        }
                        

                        //product name
                        
                        try {
                     /*       WebElement brandName = productElementInside.findElement(By.xpath(".//div[@class = 'sc-aXZVg gXaIUH _1v3Kq']"));
                            WebElement nameOfProduct = productElementInside.findElement(By.xpath(".//div[@class = 'sc-aXZVg gnOsqr _AHZN']"));

                             name1 = brandName.getText();
                             name2 = nameOfProduct.getText();

                             fullProductName = name1 + " " + name2;   */
                            // System.out.println(fullProductName);

                        	WebElement brandName = productElementInside.findElement(By.xpath(" .//div[@class= '_7I7zN']//div[@class='sc-gEvEer jlNTHb _AHZN']"));
                        	
                        	name1 = brandName.getText();
                        	
                        	fullProductName = name1;
                        	
                        	
                        	
                        } catch (NoSuchElementException e) {
                        	
                        	for(int y=0;y<500;y++) {
                        		
                        		driver.navigate().refresh();
                        		
                        		Thread.sleep(2000);
                        		
                        		driver.navigate().refresh();
                        		
                        		Thread.sleep(3000);
                            // Handle stale element reference exception by re-finding the element
                        		
                            WebElement brandName = productElementInside.findElement(By.xpath(".//div[@class='_1Cxye']/div[@class='hgjRZ']/div[2]"));
                            WebElement nameOfProduct = productElementInside.findElement(By.xpath(" .//div[@class = 'sc-gEvEer jlNTHb _AHZN']"));

                             name1 = brandName.getText();
                             System.out.println(name1);
                            name2 = nameOfProduct.getText();

                             fullProductName = name1 + " " + name2;
                        }
                        	//break;
                            // System.out.println(fullProductName);
                        } catch (StaleElementReferenceException e) {
                        	WebElement brandName = productElementInside.findElement(By.xpath(".//div[@class='_1Cxye']/div[@class='hgjRZ']/div[2]"));
                        	 name1 = brandName.getText();
                        	  System.out.println(name1);
                            e.printStackTrace();
                        }

                        
                        
                        
                        try {
                            WebElement mrp = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer hXrBXj _2XPBo _1QyO8']"));
                            String originalMrp1 = mrp.getText();
                            originalMrp = originalMrp1.replace("₹", "");
                        } catch (StaleElementReferenceException e) {
                            // Handle stale element reference exception by re-finding the element
                        	
                        	 WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                        	 
                        	 WebElement mrp = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(" .//div[@class = 'sc-gEvEer hXrBXj _2XPBo _1QyO8']")));
                        //    WebElement mrp = productElement.findElement(By.xpath(".//div[@class='sc-aXZVg gXaIUH _2XPBo _1QyO8']"));
                            String originalMrp1 = mrp.getText();
                            originalMrp = originalMrp1.replace("₹", "");
                        } catch (Exception e) {
                            e.printStackTrace();
                        }

                        
                        
                        
                        try {
                            WebElement sp = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer jlNTHb _2XPBo']"));
                            String originalSp = sp.getText();
                            spValue = originalSp.replace("₹", "");
                        } catch (StaleElementReferenceException e) {
                            // Handle stale element reference exception by re-finding the element
                            WebElement sp = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer jlNTHb _2XPBo']"));
                            String originalSp = sp.getText();
                            spValue = originalSp.replace("₹", "");
                        } catch (Exception e) {
                            e.printStackTrace();
                        }

                        
                        
                        
                        
                        try {
                            WebElement uomProduct = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer hXrBXj _1TwvP']"));
                            originalUom = uomProduct.getText();
                        } catch (StaleElementReferenceException e) {
                            // Handle stale element reference exception by re-finding the element
                            WebElement uomProduct = productElement.findElement(By.xpath(" .//div[@class = 'sc-gEvEer hXrBXj _1TwvP']"));
                            originalUom = uomProduct.getText();
                        } catch (Exception e) {
                            e.printStackTrace();
                        }

                       
                        
                    
                    Thread.sleep(2000);
                    
                    // Write data to output Excel sheet
                    Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                    outputRow.createCell(0).setCellValue(pId);
                    outputRow.createCell(1).setCellValue(city);
                    outputRow.createCell(2).setCellValue(productName);
                    outputRow.createCell(3).setCellValue(originalUom);	
                    outputRow.createCell(4).setCellValue(productUrl);
                    outputRow.createCell(5).setCellValue(fullProductName);
                    outputRow.createCell(6).setCellValue(originalMrp);
                    outputRow.createCell(7).setCellValue(spValue);
                   
                    System.out.println(productUrl);
                    System.out.println(fullProductName);
                    System.out.println(originalMrp);
                    System.out.println(spValue);
                    
                    Thread.sleep(1000); 
                   
                    
                    
                    try {
                    
                    driver.navigate().back();
                    
                    System.out.println("++++++++++  Navigate back Success    ++++++++++");
                    }
                    
                    catch(Exception t) {
                    productElements = driver.findElements(By.xpath(" .//div[@class='_13vAM']"));
                    
                    System.out.println("+++++++++=======+  Navigate back Success    =========++++++++++");
                    }
                }
                
                }
                    
                // Check if the third product was found
                if (count < 3 && !thirdProductFound) {
                    // Write "NA" to the output Excel sheet for the third product
                    Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                    outputRow.createCell(0).setCellValue(pId);
                    outputRow.createCell(1).setCellValue(city);
                    outputRow.createCell(2).setCellValue(productName);
                    outputRow.createCell(3).setCellValue(uom);
                    outputRow.createCell(4).setCellValue("NA");
                    outputRow.createCell(5).setCellValue("NA");
                    outputRow.createCell(6).setCellValue("NA");
                    outputRow.createCell(7).setCellValue("NA");
                    
                    thirdProductFound = true; // Set the flag to true since the third product was found
                }
            }

            // Save output Excel file
            FileOutputStream fileOut = new FileOutputStream(outputFile);
            outputWorkbook.write(fileOut); 
            fileOut.close();

            // Close the WebDriver
            System.out.println("DONE SCRAPING");
            driver.quit();

        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
}