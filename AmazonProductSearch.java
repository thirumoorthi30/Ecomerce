package Shopping;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class AmazonProductSearch {

    public static void main(String[] args) throws Exception {
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();

        // Load the input data from an Excel sheet
        List<String> productNames = readProductNamesFromExcel(".\\input-data\\Product Search.xlsx");

        // Create a new workbook to store the results
        Workbook resultWorkbook = new XSSFWorkbook();
        Sheet resultSheet = resultWorkbook.createSheet("Product URLs");
        int rowNum = 0;

        // Navigate to the Amazon website
        driver.get("https://www.amazon.in/?&tag=googhydrabk1-21&ref=pd_sl_7hz2t19t5c_e&adgrpid=155259815513&hvpone=&hvptwo=&hvadid=674842289437&hvpos=&hvnetw=g&hvrand=12817599141667678899&hvqmt=e&hvdev=c&hvdvcmdl=&hvlocint=&hvlocphy=9148906&hvtargid=kwd-10573980&hydadcr=14453_2316415&gad_source=1");

        // Find the search bar and search for each product
        
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
		
		
		
        
        for (String productName : productNames) {
            WebElement searchInput = driver.findElement(By.xpath("//*[@id=\"twotabsearchtextbox\"]"));
            searchInput.clear(); // Clear the search bar
            searchInput.sendKeys(productName); // Enter the product name
            searchInput.sendKeys(Keys.ENTER); // Submit the search query

            // Wait for the search results to load (you may need to adjust the wait time)
            Thread.sleep(4000);

            // Find all product elements in the search results
            List<WebElement> productElements = driver.findElements(By.xpath("//*[@data-component-type='s-search-result']"));
            
            
            

            // Limit to scrape only the first three products
            int count = Math.min(productElements.size(), 3);
            
            
            

            // Loop through the first three product elements
            List<String> productUrls = new ArrayList<>();
            for (int i = 0; i < count; i++) {
                // Get the ith product element
                WebElement productElement = productElements.get(i);

                // Find the anchor tag within the product element
                WebElement productLink = productElement.findElement(By.xpath(".//h2/a"));

                // Get the href attribute of the anchor tag
                String productUrl = productLink.getAttribute("href");

                // Add the product URL to the list
                productUrls.add(productUrl);
                
                System.out.println(productUrl);
            }

            // Write the product URLs to the Excel sheet
            writeProductUrlsToExcel(productName, productUrls, resultSheet, rowNum);
            rowNum += productUrls.size();
        }

        // Save the result workbook
        FileOutputStream outputStream = new FileOutputStream(".\\Output\\Amazon Self Serach_ Output \" + timestamp + \".xlsx");
        resultWorkbook.write(outputStream);
        resultWorkbook.close();

        // Close the browser
        driver.quit();
    }

    // Method to read product names from an Excel sheet
    private static List<String> readProductNamesFromExcel(String filePath) throws Exception {
        List<String> productNames = new ArrayList<>();
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0); // Assuming the product names are in the first sheet

        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cell = row.getCell(0); // Assuming the product names are in the first column
            if (cell != null) {
                productNames.add(cell.getStringCellValue());
            }
        }

        workbook.close();
        fis.close();

        return productNames;
    }

    // Method to write product URLs to an Excel sheet
    private static void writeProductUrlsToExcel(String productName, List<String> productUrls, Sheet sheet, int rowNum) {
        for (int i = 0; i < productUrls.size(); i++) {
            Row row = sheet.createRow(rowNum + i);
            Cell productNameCell = row.createCell(0);
            Cell productUrlCell = row.createCell(1);

            productNameCell.setCellValue(productName);
            productUrlCell.setCellValue(productUrls.get(i));
        }
    }
}
