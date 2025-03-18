package Shopping;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class PurplleProductScraper{
	public static void main(String[] args) {
        // Generate timestamp
        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String timestamp = currentTime.format(formatter);

        String inputFile = ".\\input-data\\Product Search.xlsx";
        String outputFile = ".\\Output\\Purplle_Pro_Ser_OutputData " + timestamp + " .xlsx";

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
            	 // Assuming product name is in first column 
            	String pId = row.getCell(0).getStringCellValue();
            	String city = row.getCell(1).getStringCellValue();
                String productName = row.getCell(2).getStringCellValue();
                String uom = row.getCell(3).getStringCellValue();
                driver.manage().deleteAllCookies();

                driver.get("https://www.purplle.com/");
                

                // Find the search bar and search for each product
                WebElement searchInput = driver.findElement(By.xpath("//input[@placeholder='What are you looking for?']"));
                searchInput.click();
                searchInput.clear(); // Clear the search bar
                WebElement Search=driver.findElement(By.xpath("//*[@id=\"body\"]/app-root/div/div/app-header/div[2]/desktop-search-content/div[1]/div/div[1]/p/input"));
                //Search.click();
               
                Search.sendKeys(productName); // Enter the product name
                Search.sendKeys(Keys.ENTER); // Submit the search query
                

                // Wait for the search results to load (you may need to adjust the wait time)
                Thread.sleep(4000);

                // Find all product elements in the search results
                List<WebElement> productElements = driver.findElements(By.cssSelector("div.d-block.w-100.position-relative.clearfix > app-listing-item"));

                // Limit to scrape only the first three products
                int count = Math.min(productElements.size(), 3);
                
                String productUrl = " ";
                String productname = " ";
                String originalMrp = " ";
                String spValue = " ";
                

                // Loop through the first three product elements
                for (int i = 0; i < count; i++) {
                    // Get the ith product element
                    WebElement productElement = productElements.get(i);
                    
                    
                    try {
                    // Find the anchor tag within the product element
                    WebElement productLink = productElement.findElement(By.xpath(".//a[@class='d-block mb-12p ng-star-inserted']"));
                     productUrl = productLink.getAttribute("href");
                    }
                    catch(Exception w) {
                    	productUrl="NA";
                    }
                    
                    
                    try {
                    WebElement  productNewName = productElement.findElement(By.xpath(" 	.//div[@class = 'product-title fs-7 text-start text-black fw-normal']"));
                    productname = productNewName.getText();
                    }
                    catch(Exception e) {
                    
                    }
                   
                    try {
                    WebElement mrp = productElement.findElement(By.xpath(" .//span[@class = 'text-black fw-bolder fs-6']"));
                    String originalMrp1 = mrp.getText();
                    originalMrp = originalMrp1.replace("₹","");
                  
                    
                 //  System.out.println(originalMrp);
                 //  System.out.println(originalMrp.length());
                    }
                    catch(Exception h) {
                    	originalMrp="NA";
                    }
                    
                    
                    try {
                    WebElement sp = productElement.findElement(By.xpath(" .//s[@class = 'text-black-50 ms-1 fw-medium ng-star-inserted']"));
                    String originalSp = sp.getText();
                    spValue = originalSp.replace("₹", "");
                  //  System.out.println(spValue);
                    
                    }
                    catch(Exception t) {
                    	spValue="NA";
                    }
                    
                    

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
                   // outputRow.createCell(2).setCellValue(originalOffer);
                    

                    System.out.println(productUrl);
                    System.out.println(productname);
                    System.out.println(originalMrp);
                    System.out.println(spValue);
                    //System.out.println(originalOffer);
                    
                    
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
