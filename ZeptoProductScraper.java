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
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

public class ZeptoProductScraper {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		  LocalDateTime currentTime = LocalDateTime.now();
	        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
	        String timestamp = currentTime.format(formatter);

	        String inputFile = ".\\input-data\\Product Search.xlsx";
	        String outputFile = ".\\Output\\Zepto_Pro_Ser_OutputData " + timestamp + " .xlsx";

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
	                //String productName = row.getCell(0).getStringCellValue(); // Assuming product name is in first column

	                String pId = row.getCell(0).getStringCellValue();
	                String city = row.getCell(1).getStringCellValue();
	                String productName = row.getCell(2).getStringCellValue();
	                String uom = row.getCell(3).getStringCellValue();
	                
	               driver.get("https://www.zeptonow.com");

	            //search
//	               Thread.sleep(5000);
//	               WebElement searchInput= driver.findElement(By.xpath("//div[@class='inline-block flex-1']//a"));
//	               Actions actions = new Actions(driver);
//	               actions.doubleClick(searchInput).perform();
//	               
	               //search
	               Thread.sleep(5000);
	               WebElement searchClick;
	               	boolean btn=false;
	               	
	              for(int k=0;k<200;k++) {
	            	  try {
	            		  driver.findElement(By.xpath("//a[@aria-label='Zepto Home']//following-sibling::div//a[@aria-label='Search for products']")).click();
	            		  break;
	            	  } catch (Exception e) {
						if (k==199) {
							Assert.fail(e.getMessage());
						}
					}
	             
	               
	              }
	              for (int i = 0; i < 200; i++) {
					try {
						driver.findElement(By.xpath("//div[@id='__next']//input")).sendKeys(productName,Keys.ENTER);
						break;
					} catch (Exception e) {
						if (i==199) {
							Assert.fail(e.getMessage());
						}
					}
				}
	              
	               Thread.sleep(2000);
	               
	               
	              // List<WebElement> productElements = driver.findElements(By.xpath("//div[contains(@class, 'items-list-container_search-item-container__Jkq1Q')]"));
	               WebDriverWait wai = new WebDriverWait(driver,  Duration.ofSeconds(10));
	             //  List<WebElement> productElements = wai.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//div[contains(@class, 'items-list-container_search-item-container__Jkq1Q')]")));
	            //   System.out.println(productElement);
	             //  Thread.sleep(5000);
	              // int count = Math.min(productElements.size(), 3);
	             
//	            	   for (int i = 0; i < count; i++) {
//	            		    WebElement productElement = productElements.get(i);
		             //  wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[@class=\"space-y-5\"]"))); // Find the top three search results 
		               
		               List<WebElement> productElements = driver.findElements(By.xpath("//div[@id='__next']//div[contains(@class,'list-container')]//a"));
		               System.out.println(productElements);
	               // Find the top three search results List<WebElement> searchResults = driver.findElements(By.cssSelector("div.product-list > div.product"));
	               int count = Math.min(productElements.size(), 3);
	             
	               			
            	   for (int i = 0; i <count; i++) {
           		    WebElement productElement = productElements.get(i);
	            	   
	                String productUrl = " ";
	                String productname = " ";
	                String originalMrp = " ";
	                String spValue = " ";
	                
	                try {
                        // Find the anchor tag within the product element
                    //    WebElement productLink = productElement.findElement(By.xpath(".//a"));
                        productUrl = productElement.getAttribute("href");
                    }
                    catch(Exception w) {
                    	
                    }
//	                for (int i = 0; i < count; i++) {
//	                    // Get the ith product element
//	                    WebElement productElement = productElements.get(i);
//	                    
//	                    try {
//	                        // Find the anchor tag within the product element
//	                        WebElement productLink = productElement.findElement(By.xpath(".//a"));
//	                        productUrl = productLink.getAttribute("href");
//	                    }
//	                    catch(Exception w) {
//	                    	
//	                    }
	                    try {
	                        WebElement  productNewName = productElement.findElement(By.xpath("//div[contains(@class,'product-card_footer')]/preceding-sibling::div//h4[contains(@class,'clamp-2')]"));
	                        productname = productNewName.getText();
	                    }
	                    catch(Exception e) {
	                    
	                    }
	                    
	                   
	                    try {
	                        WebElement mrp = productElement.findElement(By.xpath("//div[contains(@class,'product-card_footer')]//p"));
	                        String originalMrp1 = mrp.getText();
	                        originalMrp = originalMrp1.replace("₹","");
	                    }
	                    catch(Exception h) {
	                    	
	                    }
	                    
	                    
	                    try {
	                        WebElement sp = productElement.findElement(By.xpath("div[contains(@class,'product-card_footer')]//h4"));
	                        String originalSp = sp.getText();
	                        spValue = originalSp.replace("₹", "");
	                    }
	                    catch(Exception t) {
	                    	
	                    }
	                    
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
	                    //System.out.println(originalOffer);
	               	 
	                }
	            }
	            
	            FileOutputStream fileOut = new FileOutputStream(outputFile);
	  		  outputWorkbook.write(fileOut); 
	  		  fileOut.close();
	  		 

	          // Close the WebDriver
	          driver.quit();

	      } catch (IOException | InterruptedException e) {
	          e.printStackTrace();
	      }
	        catch (Exception e) {
				
			}

	        
	  }
	}


