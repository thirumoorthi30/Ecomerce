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

public class SwiggyProductLinkScraper {

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
              //  List<WebElement> productElements = driver.findElements(By.xpath(" .//div[@class='XjYJe']"));
                
                

            
                String productUrl = " ";
                String name1 = " ";
                String name2 = " ";
                String fullProductName = " ";
                String originalMrp = " ";
                String spValue = " ";
                String originalUom = " ";
                
                // Flag to check if the third product was found
                boolean thirdProductFound = false;
                
               // int count = Math.min(productElements.size(), 3);
                
                List<WebElement> productElements = driver.findElements(By.xpath(" .//div[@class='XjYJe']"));
                
                System.out.println(productElements);
                
               // System.out.println(count);
                
                // Loop through the first three product elements
                for (WebElement productElement : productElements) {
                    // Get the first three product elements within the search result
                //    List<WebElement> productElementsInside = productElement.findElements(By.xpath(".//div[@class='XjYJe']"));

              /*      // Get the ith product element
                	 WebElement productElement = null;
                	    try {
                	        productElement = productElements.get(i);
                	    } catch (IndexOutOfBoundsException e) {
                	        
                	        System.err.println("Index out of bounds: " + e.getMessage());
                	        continue; 
                	    }
          */
                    Thread.sleep(2000);
                    
                    List<WebElement> productElementsList = driver.findElements(By.xpath(" .//div[@class='_1OklP _1TqJT']"));
                    
                 //   List<WebElement> productElements1 = driver.findElements(By.xpath(" .//div[@class='XjYJe']"));
                 // Loop through each product element in the list
                 // Loop through each product element in the list
                    for (int i = 0; i < Math.min(productElementsList.size(), 4); i++) {
                        WebElement productElementInside = productElementsList.get(i);
                        try {
                            // Click on the product element
                            productElementInside.click();
                            
                            // Add a short delay to allow the page to load after clicking
                            Thread.sleep(2000);
                            
                            // Retrieve the product URL
                            productUrl = driver.getCurrentUrl();
                            System.out.println("Product URL: " + productUrl);
                            
                            // Retrieve other information about the product and perform necessary actions
                            
                            // Go back to the search results page
                            driver.navigate().back();
                            
                            // Add a short delay before processing the next product
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
                        } catch (StaleElementReferenceException e) {
                            // Handle the stale element exception by refreshing the list of product elements
                            productElementsList = driver.findElements(By.xpath(" .//div[@class='_1OklP _1TqJT']"));
                        } catch (Exception e) {
                            e.printStackTrace();
                            System.out.println("Failed to process product.");
                        }
                    }
break;
                
                }
                    
                // Check if the third product was found
             /*   if (count < 3 && !thirdProductFound) {
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
                }   */
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


