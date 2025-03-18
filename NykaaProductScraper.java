package Shopping;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
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

public class NykaaProductScraper {

    public static void main(String[] args) {
        // Generate timestamp
        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String timestamp = currentTime.format(formatter);

        String inputFile = ".\\input-data\\Product Search.xlsx";
        String outputFile = ".\\Output\\Nykaa_Pro_Ser_OutputData " + timestamp + " .xlsx";

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

                driver.get("https://www.nykaa.com/");
                
                driver.manage().deleteAllCookies();

                // Find the search bar and search for each product
                WebElement searchInput = driver.findElement(By.xpath("//*[@id=\"headerMenu\"]/div[2]/div/div/form/input"));
                searchInput.clear(); // Clear the search bar
                searchInput.sendKeys(productName); // Enter the product name
                searchInput.sendKeys(Keys.ENTER); // Submit the search query
                
                driver.manage().deleteAllCookies();

                // Wait for the search results to load (you may need to adjust the wait time)
                Thread.sleep(4000);

                // Find all product elements in the search results
                List<WebElement> productElements = driver.findElements(By.xpath("//*[@class = 'css-d5z3ro']"));

                // Limit to scrape only the first three products
                int count = Math.min(productElements.size(), 3);
                
                String productUrl = " ";
                String productname = " ";
                String originalMrp = " ";
                String spValue = " ";
                
                // Flag to check if the third product was found
                boolean thirdProductFound = false;

                // Loop through the first three product elements
                for (int i = 0; i < count; i++) {
                    // Get the ith product element
                    WebElement productElement = productElements.get(i);
                    
                    driver.manage().deleteAllCookies();
                    
                    try {
                        // Find the anchor tag within the product element
                        WebElement productLink = productElement.findElement(By.xpath(".//a[@class='css-qlopj4']"));
                        productUrl = productLink.getAttribute("href");
                    }
                    catch(Exception w) {
                    	
                    }
                    
                    try {
                        WebElement  productNewName = productElement.findElement(By.xpath(" .//div[@class = 'css-xrzmfa']"));
                        productname = productNewName.getText();
                    }
                    catch(Exception e) {
                    
                    }
                    
                   
                    try {
                        WebElement mrp = productElement.findElement(By.xpath(" .//span[@class = 'css-17x46n5']"));
                        String originalMrp1 = mrp.getText();
                        originalMrp = originalMrp1.replace("MRP:₹","");
                    }
                    catch(Exception h) {
                    	
                    }
                    
                    
                    try {
                        WebElement sp = productElement.findElement(By.xpath(" .//span[@class = 'css-111z9ua']"));
                        String originalSp = sp.getText();
                        spValue = originalSp.replace("₹", "");
                    }
                    catch(Exception t) {
                    	
                    }
                    
                    driver.manage().deleteAllCookies();
                    
                    // Write data to output Excel sheet
                    Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                    outputRow.createCell(0).setCellValue(pId);
                    outputRow.createCell(1).setCellValue(city);
                    outputRow.createCell(2).setCellValue(productName);
                    outputRow.createCell(3).setCellValue(uom);
                    outputRow.createCell(4).setCellValue(productUrl);
                    outputRow.createCell(5).setCellValue(productname);
                    outputRow.createCell(6).setCellValue(originalMrp);
                    outputRow.createCell(7).setCellValue(spValue);
                   
                    System.out.println(productUrl);
                    System.out.println(productname);
                    System.out.println(originalMrp);
                    System.out.println(spValue);
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
