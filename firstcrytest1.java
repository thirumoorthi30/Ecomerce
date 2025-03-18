package Dailyrun;
import org.openqa.selenium.chrome.ChromeDriver;
import CommonUtility.BlinkitId;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class firstcrytest1 {

	public static void main(String[] args) throws Exception{
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        int count = 0;
        // int finalSp;
          String spValue = "";
          String finalSp = "";
          String newName = null;
          String mrpValue = null;
          String offerValue =null;
          String OfferCode =null;
          String DcountValue=null;
          String dcountCode=null;
          String Offer1=null;
          String sp1=null;
          String Discount1=null;
          String sp2=null;
          
          
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\firstcryDiaHyd.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Sheet4");
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
            headerRow.createCell(9).setCellValue("OFFER 1");
            headerRow.createCell(10).setCellValue("SP 1");
            headerRow.createCell(11).setCellValue("DISCOUNT 1");
            headerRow.createCell(12).setCellValue("SP 2");
      
            
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
                          resultRow.createCell(9).setCellValue(Offer1);
                          resultRow.createCell(10).setCellValue(sp1);
                          resultRow.createCell(11).setCellValue(Discount1);
                          resultRow.createCell(12).setCellValue(sp2);
                         
                          
                          System.out.println("Skipped processing for URL: " + url);
                          continue; // Skip to the next iteration
                      }
                	  

                      if(i == 0 ) {
                     	 driver.get("https://www.firstcry.com");
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
                    clickLocation.sendKeys("110001");
                    
                    WebElement clickApply = driver.findElement(By.xpath("//*[@id=\"epincode\"]/div"));
                    clickApply.click();
                    
                    }
                    Thread.sleep(5000);
                    WebElement reg=driver.findElement(By.xpath("/html/body/div[1]/div[5]/div/div[3]/ul/li[7]"));
                    reg.click();
                    Thread.sleep(5000);
                    
                    
                    WebElement regmail=driver.findElement(By.xpath("//*[@id=\"lemail\"]"));
                    regmail.click();
                    Thread.sleep(5000);
                    regmail.sendKeys("blinkittestautomation01@gmail.com");
                    
                    WebElement conmail= driver.findElement(By.xpath("//*[@id=\"login\"]/div/div[3]/span"));
                    conmail.click();
                    
                    Thread.sleep(20000);
                      }
                      
                      driver.get(url);
                    
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
                    boolean present=false;
                    try {
                     WebElement offer = driver.findElement(By.xpath("//span[@class='nn-club-tx']"));
                     offerValue = offer.getText();
                     present=offerValue.contains("Extra");
                    }catch (NoSuchElementException e) {
                   	 Offer1="NA";
                  	     sp1="NA";
					}
                     
                    if(present == true ){
                      	 
                    	 WebElement offerxpath = driver.findElement(By.xpath("//div[@class='J13SB_42 cl_fff cpn_code_box bg_29']"));
                    	OfferCode = offerxpath.getText();
                    	System.out.println(OfferCode);
                    	Offer1=offerValue.substring(6, 12);
                  	  }
                    
                    try {
                    	Thread.sleep(500);
                    	//for (int j = 0; j < 150; j++) {
                    	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                    	driver.findElement(By.xpath("(//span[@class='step1 M16_white'])[1]//span")).click();
                    	Thread.sleep(1000);
                    	driver.findElement(By.xpath("(//span[@class='step2 M16_white'])[1]")).click();
                    	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                    	// break;
                    	// }
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
                   
                    WebElement rate = driver.findElement(By.id("fnl_pymnt"));
                    rateValue = rate.getText();
                    double paiseDecimal = Double.parseDouble(rateValue) / 100;
                    System.out.println(paiseDecimal);
                    rateValue = String.valueOf(paiseDecimal);

                    finalSp = rateValue;
                    
                    if(present=true) {
                        Thread.sleep(5000);
                       WebElement coupon=driver.findElement(By.xpath("/html/body/form/section[1]/section[4]/div[4]/div[8]/div/div[3]/div/input"));
                       coupon.click();
                       coupon.clear();
                       coupon.sendKeys(OfferCode);
                       
                       Thread.sleep(1000);
                       WebElement aplcoupon=driver.findElement(By.xpath("/html/body/form/section[1]/section[4]/div[4]/div[8]/div/div[3]/div/span[2]"));
                       aplcoupon.click();
                       
                       Thread.sleep(2000);
                       
                       String OFFrateValue="";
                       WebElement OFFrate = driver.findElement(By.id("fnl_pymnt"));
                       OFFrateValue = OFFrate.getText();
                       double paisedecimal = Double.parseDouble(OFFrateValue) / 100;
                       System.out.println(paisedecimal);
                       OFFrateValue = String.valueOf(paisedecimal);
 }
                        
                       
                  
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
                    resultRow.createCell(9).setCellValue(Offer1);
                    resultRow.createCell(10).setCellValue(sp1);
                    resultRow.createCell(11).setCellValue(Discount1);
                    resultRow.createCell(12).setCellValue(sp2);
                    
                    
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
                    resultRow.createCell(9).setCellValue(Offer1);
                    resultRow.createCell(10).setCellValue(sp1);
                    resultRow.createCell(11).setCellValue(Discount1);
                    resultRow.createCell(12).setCellValue(sp2);
                   

                    System.out.println("Failed to extract data for URL: " + url);
                    
                }
            }
            
            // Write results to Excel file
            FileOutputStream outFile = new FileOutputStream(".\\Output\\Firstcry Dia outputData.xlsx");
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
