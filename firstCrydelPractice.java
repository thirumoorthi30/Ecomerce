package scheduler;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import CommonUtility.BlinkitId;

public class firstCrydelPractice {
	public static void main(String[] args) {
		ChromeOptions options = new ChromeOptions();
		options.addArguments("--headless"); // Run Chrome in headless mode
		options.addArguments("--disable-gpu"); // Disable GPU acceleration

		ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);

		// Schedule the task to run every day at 7:00 AM
		Calendar now = Calendar.getInstance();
		Calendar nextRunTime = Calendar.getInstance();
		nextRunTime.set(Calendar.HOUR_OF_DAY, 10);
		nextRunTime.set(Calendar.MINUTE, 49);
		nextRunTime.set(Calendar.SECOND, 0);

		long initialDelay = nextRunTime.getTimeInMillis() - now.getTimeInMillis();
		if (initialDelay < 0) {
		initialDelay += 24 * 60 * 60 * 1000; // If it's already past 7 AM, schedule for the next day
		}

		scheduler.scheduleAtFixedRate(() -> {
		try {
		System.out.println("Starting web scraping task...");
		firstCrydelPractice.runWebScraping();
		System.out.println("Web scraping task completed.");
		} catch (Exception e) {
		e.printStackTrace();
		}
		}, initialDelay, 24 * 60 * 60 * 1000, TimeUnit.MILLISECONDS);
		}
	
	public static void runWebScraping() throws Exception{
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        int count = 0;
        // int finalSp;
          String spValue = "";
          String finalSp = "";
          String newName = null;
          String mrpValue = null;
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\firstcryDiaDel.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Sheet3");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

	            List<String> inputPid = new ArrayList<>(),InputCity = new ArrayList<>(),InputName = new ArrayList<>(),InputSize = new ArrayList<>(),NewProductCode = new ArrayList<>(),
	            		uRL = new ArrayList<>();
	            
            // Extract URLs from Excel
            for (int i = 0; i < rowCount; i++) {
                Row row = urlsSheet.getRow(i);
                
                
                  if (i == 0) {
                continue;
            }    
                
                
                Cell inputPidCell = row.getCell(0);
                Cell inputCityCell = row.getCell(1);
                Cell inputNameCell = row.getCell(2);
                Cell inputSizeCell = row.getCell(3);
                Cell newProductCodeCell = row.getCell(4);
                Cell urlCell = row.getCell(5);
            //    Cell urlCell = row.getCell(0);
              //  Cell urlCell = row.getCell(0);
               // Cell idCell = row.getCell(1);
               // Cell offerCell = row.getCell(2);
                
                if (urlCell != null && urlCell.getCellType() == CellType.STRING) {
                    String url = urlCell.getStringCellValue();
                    String id = (inputPidCell != null && inputPidCell.getCellType() == CellType.STRING) ? inputPidCell.getStringCellValue() : "";
                    String city = (inputCityCell != null && inputCityCell.getCellType() == CellType.STRING) ? inputCityCell.getStringCellValue() : "";
                    String name = (inputNameCell != null && inputNameCell.getCellType() == CellType.STRING) ? inputNameCell.getStringCellValue() : "";
                    String size = (inputSizeCell != null && inputSizeCell.getCellType() == CellType.STRING) ? inputSizeCell.getStringCellValue() : "";
                    String productCode = (newProductCodeCell != null && newProductCodeCell.getCellType() == CellType.STRING) ? newProductCodeCell.getStringCellValue() : "";
                    
                    inputPid.add(id);
                    InputCity.add(city);
                    InputName.add(name);
                    InputSize.add(size);
                    NewProductCode.add(productCode);
                    uRL.add(url);
					/*
					 * uRL.add(url); ids.add(id); offers.add(offer);
					 */
                    
                }
            }
            // Create Excel workbook for storing results
            Workbook resultsWorkbook = new XSSFWorkbook();
            Sheet resultsSheet = resultsWorkbook.createSheet("Results");

            Row headerRow = resultsSheet.createRow(0);
            
            
            headerRow.createCell(0).setCellValue("InputPid");
            headerRow.createCell(1).setCellValue("InputCity");
            headerRow.createCell(2).setCellValue("InputName");
            headerRow.createCell(3).setCellValue("InputSize");
            headerRow.createCell(4).setCellValue("NewProductCode");
            headerRow.createCell(5).setCellValue("URL");
            headerRow.createCell(6).setCellValue("Name");
            headerRow.createCell(7).setCellValue("MRP");
            headerRow.createCell(8).setCellValue("SP");
            
            int rowIndex = 1;

            int headercount = 0;
            
            for (int i = 0; i < uRL.size(); i++) {
                String id = inputPid.get(i);
                String city = InputCity.get(i);
                String name = InputName.get(i);
                String size = InputSize.get(i);
                String productCode = NewProductCode.get(i);
                String url = uRL.get(i);
                
                try {
                	
                	  if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
                          // Set "NA" values in all three columns
                          Row resultRow = resultsSheet.createRow(rowIndex++);
                          resultRow.createCell(0).setCellValue(id);
                          resultRow.createCell(1).setCellValue(city);
                          resultRow.createCell(2).setCellValue(name);
                          resultRow.createCell(3).setCellValue(size);
                          resultRow.createCell(4).setCellValue(productCode);
                          resultRow.createCell(5).setCellValue(url);
                          resultRow.createCell(6).setCellValue("NA");
                          resultRow.createCell(7).setCellValue("NA");
                          resultRow.createCell(8).setCellValue("NA");
                          
                          System.out.println("Skipped processing for URL: " + url);
                          continue; // Skip to the next iteration
                      }
                	  
                    driver.get(url);
                    driver.manage().window().maximize();
                    
                    //location set
                    if(i == 0 ||i == 1) {

                    Thread.sleep(1000);
                    WebElement locationSet= driver.findElement(By.xpath("//*[@id=\"geoLocation\"]/span/div[1]/span"));
                    locationSet.click();
                    Thread.sleep(1000);
                    
                    WebElement setLocation = driver.findElement(By.xpath("//*[@id=\"pincodetext\"]/div/sapn"));
                    setLocation.click();
                    
                    WebElement clickLocation = driver.findElement(By.xpath("//*[@id=\"nonlpincode\"]"));
                    clickLocation.clear();
                    clickLocation.sendKeys("500001");
                    
                    WebElement clickApply = driver.findElement(By.xpath("//*[@id=\"epincode\"]/div"));
                    clickApply.click();
                    
                    }
                    
                    
                    try {
                    	
                    WebElement nameElement = driver.findElement(By.id("prod_name"));
                    newName = nameElement.getText();
                    System.out.println(newName);
                    }
                    
                    catch(org.openqa.selenium.NoSuchElementException e) {
                    	
                    	// WebElement nameElement = driver.findElement(By.id("prod_name"));
                    	
                    	WebElement nameElement = driver.findElement(By.xpath("//div[@class = 'prod-info-wrap']//following::p[1]"));
                    	newName = nameElement.getText();
                        System.out.println(newName);
                    	
                    }
                    
                    System.out.println("headercount = " + headercount);
                    
                    headercount++;
                    
                    
                    try {
                    WebElement mrp = driver.findElement(By.xpath("//*[@id=\"original_mrp\"]"));
                    mrpValue = mrp.getText();
                    System.out.println(mrpValue);
                    
                    }
                    
                    catch(org.openqa.selenium.NoSuchElementException e){
                    	
                        WebElement mrp = driver.findElement(By.xpath("/html/body/div[5]/div/div[2]/div[2]/div[2]/div[2]/span[4]/span[3]"));
                        mrpValue = mrp.getText();                            
                        System.out.println(mrpValue);
                    	
                    }
                    
                    try {
                    Thread.sleep(500);
                    
                    //for (int j = 0; j < 150; j++) {
                       //  driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                     	driver.findElement(By.xpath("(//span[@class='step1 M16_white'])[1]//span")).click();
                 		
                     	 Thread.sleep(1000); 
                     	
                 		driver.findElement(By.xpath("(//span[@class='step2 M16_white'])[1]")).click();
                 		
                        // driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                        // break;
                   //  }
                    }
                    catch(NoSuchElementException e){
                    	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                    	
                    	try {
                    		Thread.sleep(500);
                        	driver.findElement(By.xpath("//*[@id=\"p_breadcrumb\"]/div[2]/div/div[2]/div[7]/div/div[2]/div[2]/div[1]/div/span[1]/span")).click();
                     		
                        	 Thread.sleep(2000); 
                        	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                    		driver.findElement(By.xpath("//*[@id=\"p_breadcrumb\"]/div[2]/div/div[2]/div[7]/div/div[2]/div[2]/div[1]/div/span[2]")).click();
                    		
                    	}
                    	catch(NoSuchElementException ex) {
                    		Thread.sleep(500);
                        	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                        	driver.findElement(By.xpath("/html/body/div[5]/div[2]/div/div[2]/div[7]/div/div[2]/div[2]/div[1]/div/span[1]/span")).click();
                     		
                        	 Thread.sleep(2000); 
                        	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                    		driver.findElement(By.xpath("/html/body/div[5]/div[2]/div/div[2]/div[7]/div/div[2]/div[2]/div[1]/div/span[2]")).click();
                    		
                    	}
                		
                    }
                    //catch the 2 product has the different xpath for add to card
                    catch(Exception ex){
                    	
                    	//count = 1;
                    	
                    	WebElement nameElement = driver.findElement(By.id("prod_name"));
                        newName = nameElement.getText();
                        
                        Thread.sleep(500);
                        
                        WebElement mrp = driver.findElement(By.xpath("//*[@id=\"original_mrp\"]"));
                        mrpValue = mrp.getText();
                        
                        Thread.sleep(1000);
                        
                        WebElement sp = driver.findElement(By.xpath("//*[@id=\"prod_price\"]"));
                       spValue = sp.getText();
                       System.out.println("+++++++++++" + spValue);
                        finalSp = spValue;
                        
                       count=1;
                    }
                    System.out.println("==================="+count+"===================");
                    
                    
                    BlinkitId screenshot = new BlinkitId();

                    try {
        				screenshot.screenshot(driver, "Firstcry", id);
        			} catch (Exception e) {
        				e.fillInStackTrace();
        			
        			}
                    
                    
                    finalSp = spValue;
                    
                    
                    System.out.println("==================="+finalSp+"===================");
                    
                  //  finalSp = Integer.parseInt(spValue);
                    
                    String rateValue = "";
                    
                  //  if(count != 1 ) {
                   
                    WebElement rate = driver.findElement(By.className("fnl_pymnt"));
                    rateValue = rate.getText();
                    System.out.println(rateValue);
                    finalSp = rateValue;
               //     }
                    //rateValue = spValue;
                   
                    
                    
                    //  boolean success = false;
                      //  int attempts = 0;
                        	//if(count != 1) {
                    //    while (!success && attempts < 3) { 
                            try {
                            	Thread.sleep(4000); 
                                //driver.findElement(By.xpath("//*[@id=\"garem_3312344\"]")).click();
                                
                             WebElement remove = driver.findElement(By.xpath("//*[@id=\"productlist\"]/div/div[6]/div[1]/span[2]"));
                           //*[@id="productlist"]/div/div[3]/div[1]/span[2]
                            // Thread.sleep(1000);
                             
                             remove.click();

                              //  success = true;
                            } catch (Exception e) {
                               
                            	Thread.sleep(4000); 
                                //driver.findElement(By.xpath("//*[@id=\"garem_3312344\"]")).click();
                                
                             WebElement remove = driver.findElement(By.xpath("//*[@id=\"productlist\"]/div/div[3]/div[1]/span[2]"));
                           //*[@id="productlist"]/div/div[3]/div[1]/span[2]
                           //  Thread.sleep(1000);
                             
                             remove.click();

                                e.printStackTrace();
                               // attempts++;
                                Thread.sleep(2000); 
                            }
                            
                     //   }  
                        	//}
                    
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    
                    resultRow.createCell(0).setCellValue(id);
                    resultRow.createCell(1).setCellValue(city);
                    resultRow.createCell(2).setCellValue(name); 
                    resultRow.createCell(3).setCellValue(size);
                    resultRow.createCell(4).setCellValue(productCode);
                    resultRow.createCell(5).setCellValue(url);
                    resultRow.createCell(6).setCellValue(newName);
                    resultRow.createCell(7).setCellValue(mrpValue);
                    resultRow.createCell(8).setCellValue(finalSp);
                    
                    System.out.println("Data extracted for URL: " + url);
                } catch (Exception e) {
                    e.printStackTrace();
                    
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    resultRow.createCell(0).setCellValue(id);
                    resultRow.createCell(1).setCellValue(city);
                    resultRow.createCell(2).setCellValue(name);
                    resultRow.createCell(3).setCellValue(size);
                    resultRow.createCell(4).setCellValue(productCode);
                    resultRow.createCell(5).setCellValue(url);
                    resultRow.createCell(6).setCellValue("NA");
                    resultRow.createCell(7).setCellValue("NA");
                    resultRow.createCell(8).setCellValue("NA");

                    System.out.println("Failed to extract data for URL: " + url);
                    
                }
            }
            
            // Write results to Excel file
            FileOutputStream outFile = new FileOutputStream(".\\Output\\Firstcry Dia OutputData_FirstHalf.xlsx");
            resultsWorkbook.write(outFile);
            outFile.close();

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
                driver.quit();
            }
        }
    }

}
