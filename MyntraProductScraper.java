package Shopping;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
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

public class MyntraProductScraper {

    public static void main(String[] args) {
        // Generate timestamp
        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String timestamp = currentTime.format(formatter);

        String inputFile = ".\\input-data\\Product Search.xlsx";
        String outputFile = ".\\Output\\Myntra_Pro_Ser_OutputData " + timestamp + " .xlsx";

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
            headerRow.createCell(0).setCellValue("PID");
            headerRow.createCell(1).setCellValue("CITY");
            headerRow.createCell(2).setCellValue("Input Product Name");
            headerRow.createCell(3).setCellValue("UOM");
            headerRow.createCell(4).setCellValue("Product URL");
            headerRow.createCell(5).setCellValue("Product Name");
            headerRow.createCell(6).setCellValue("MRP");
            headerRow.createCell(7).setCellValue("SP");
            headerRow.createCell(8).setCellValue("uom");
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
                
                driver.manage().deleteAllCookies();

                driver.get("https://www.myntra.com");
                
                // Find the search bar and search for each product
                WebElement searchInput = driver.findElement(By.xpath("//input[@class='desktop-searchBar']"));
                searchInput.clear(); // Clear the search bar
                searchInput.sendKeys(productName); // Enter the product name
                searchInput.sendKeys(Keys.ENTER); // Submit the search query

                // Wait for the search results to load (you may need to adjust the wait time)
                Thread.sleep(4000);

                
                String productUrl = " ";
                String productname1 = " ";
                String productname2 = " ";
                String originalMrp = " ";
                String spValue = " ";
                String mrpvalue =" ";
                String fullName =" ";

                // Find all product elements in the search results
                try {
                WebDriverWait wai = new WebDriverWait(driver, Duration.ofSeconds(10));
                WebElement productElements = wai.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("ul.results-base")));
                List<WebElement> listItems = productElements.findElements(By.cssSelector("li.product-base"));

                // Initialize flag to track if any product found
                boolean productFound = false;
                
                
//                String productUrl = " ";
//                String productname1 = " ";
//                String productname2 = " ";
//                String originalMrp = " ";
//                String spValue = " ";
//                String mrpvalue =" ";
//                String fullName =" ";

                // Loop through the first three product elements
                int count = 0;
                for (WebElement listItem : listItems) {
                    if (count >= 3) {
                        break;
                    }

                    // Get the ith product element
                    try {
                        // Find the anchor tag within the product element
                        WebElement productLink = listItem.findElement(By.xpath(".//a"));
                        productUrl = productLink.getAttribute("href");
                        productFound = true; // Set flag to true if product found
                    } catch (Exception w) {
                        // Handle exception if product link not found
                    }
                    try {
//                      WebElement  productNewName1 = listItem.findElement(By.xpath("//div[@class='product-productMetaInfo']//h3"));
//                      WebElement  productNewName2 = listItem.findElement(By.xpath("//div[@class='product-productMetaInfo']//h4"));
//                      
                 	 WebElement productNewName1=listItem.findElement(By.cssSelector("h3.product-brand"));
                 	 WebElement productNewName2=listItem.findElement(By.cssSelector("h4.product-product"));
                 	 
                      productname1 = productNewName1.getText();
                      productname2 = productNewName2.getText();
                      
                      fullName = productname1 + " "+ productname2;
                      System.out.println(fullName);
                      
                      }
                      catch(Exception e) {
                      
                      }
                     
                  try {
                      WebElement mrp = listItem.findElement(By.cssSelector("span.product-strike"));
                      originalMrp = mrp.getText();
                      mrpvalue =originalMrp.replace("Rs.", "");
                     System.out.println(mrpvalue);
                  
                      }
                      catch(Exception h) {
                      	WebElement mrp = listItem.findElement(By.cssSelector("div.product-price"));
                      	originalMrp = mrp.getText();
                         mrpvalue =originalMrp.replace("Rs.", "");
                      }
                  try {
                      WebElement sp = listItem.findElement(By.cssSelector("span.product-discountedPrice"));
                      String originalSp = sp.getText();
                      spValue = originalSp.replace("Rs.", "");
                      System.out.println(spValue);
                      
                      }
                      catch(Exception t) {
                      	//spValue = mrpvalue;
                      }
                    
                }

                // Check if no product was found
                if (!productFound) {
                    for (int i = 0; i < 3; i++) {
                        // Write "NA" to the output Excel sheet for each product
                        Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                        outputRow.createCell(0).setCellValue(pId);
                        outputRow.createCell(1).setCellValue(city);
                        outputRow.createCell(2).setCellValue(productName);
                        outputRow.createCell(3).setCellValue(uom);
                        outputRow.createCell(4).setCellValue("NA");
                        outputRow.createCell(5).setCellValue("NA");
                        outputRow.createCell(6).setCellValue("NA");
                        outputRow.createCell(7).setCellValue("NA");
                    }
                }
            }catch (NoSuchElementException e) {
            	productUrl="NA";
            	fullName="NA";
            	mrpvalue="NA";
            	spValue="NA";
            	
            	
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
