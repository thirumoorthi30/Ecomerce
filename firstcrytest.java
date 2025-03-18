package Shopping;

import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.graphbuilder.math.Expression;

import CommonUtility.BlinkitId;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class firstcrytest {

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
          String Offer1=null;
          String sp1=null;
          
          
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\firstcryDiaDel.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Sheet5");
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
                          resultRow.createCell(6).setCellValue(newName);
                          resultRow.createCell(7).setCellValue(mrpValue);
                          resultRow.createCell(8).setCellValue(finalSp);
                          resultRow.createCell(9).setCellValue(Offer1);
                          resultRow.createCell(10).setCellValue(sp1);
                          
                          System.out.println("Skipped processing for URL: " + url);
                          continue; // Skip to the next iteration
                      }
                	 
                     if(i == 0 ) {
                    	 driver.get("https://www.firstcry.com");
                         driver.manage().window().maximize();
                         
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
                      
                      
//                      Thread.sleep(5000);
//                      WebElement reg=driver.findElement(By.xpath("/html/body/div[1]/div[5]/div/div[3]/ul/li[7]"));
//                      reg.click();
//                      Thread.sleep(5000);
//                      
//                      
//                      WebElement regmail=driver.findElement(By.xpath("//*[@id=\"lemail\"]"));
//                      regmail.click();
//                      Thread.sleep(5000);
//                      regmail.sendKeys("blinkitproofofconcept@gmail.com");
//                      
//                      WebElement conmail= driver.findElement(By.xpath("//*[@id=\"login\"]/div/div[3]/span"));
//                      conmail.click();
//                      
//                      
//                      Thread.sleep(20000);
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
                      
                      List<WebElement> divElements = driver.findElements(By.xpath("//div[@class='cpn_offrs_disc_section']//div[@class='swiper-slide']"));
                      Map<String, String> keyValueMap = new HashMap<>();
                      int divCount = divElements.size();
                      System.out.println("Number of <div> elements on the page: " + divCount);
                      for(int j=1;j<=divCount;j++) {
                    	 // for(int k=1;k<=divCount;k++) {
                    	  String text = null;
                    	  String discountValue=null;
                    	  String discountCode=null;
                    	  String offerCodeFlat = null;
                    	  String offerCode=null;
                    	  try {
                    		  WebElement OffervalueText = driver.findElement(By.xpath("//div[@class='cpn_offrs_disc_section']//div[@class='swiper-slide']["+j+"]//div[@class='save_cpn_header']//span[@class='nn-club-tx']"));
                              text = OffervalueText.getText();
                             System.out.println("Text : "+text);
                    		 
                             if (OffervalueText.isDisplayed()) {
                            	    System.out.println("Element is displayed on the webpage.");
                            	    // Perform actions when the element is visible
                            	} else {
                            	    System.out.println("Element is not displayed on the webpage.");
                            	    // Perform alternative actions or handle the situation
                            	}

                      WebElement OffervalueTextr = driver.findElement(By.xpath("(//span[@class='nn-club-tx'])[" + j + "]"));
                       text = OffervalueTextr.getText();
                      System.out.println("Text : "+OffervalueTextr);
                     
                      if(text.contains("All Users")) {
             			 
             			 String[] parts = text.split("All Users - ");

             		        if (parts.length >= 2) {
             		            // Extracting the discount portion
             		            String discountText = parts[1]; // This will be "Extra 5% Off* on Diapers"

             		            // Further splitting to get just the discount value
             		            String[] discountParts = discountText.split(" ");
             		            if (discountParts.length >= 2) {
             		                 discountValue = discountParts[1]; // This will be "5%"
             		               //System.out.println("Discount value: " + discountValue);
             		            }
             		 }}
             		 if(text.contains("All Users")) {
             			 int startIndex = text.indexOf("5% Off");

             		        // Extract "5% Off" from the offer text
             		         discountValue = text.substring(startIndex, startIndex + 6).trim(); // Adjust the length if "5% Off" changes

             		        // Print the extracted discount value
             		        System.out.println("Discount value: " + discountValue);}
                      }
                      catch(Exception e) {

//                    	  WebElement Offertext1=driver.findElement(By.xpath("(//div[@class='club-txt-logo'])["+j+"]"));
//                    	  String Maintext =Offertext1.getText();
//                    	  System.out.println(Maintext);
                    	  
                    	  
                      }
                    	  

                     
                      if((text.contains("Extra") && text.contains("All Users"))||text.contains("extra")) {
                    	  //div[@class='cpncode_details_row']//div[@class='cpn_code_copy_block']//div
                    	 // WebElement code = driver.findElement(By.xpath("//div[@class='cpncode_details_row']//div[@class='cpn_code_copy_block']//div"));
                    	  WebElement code = driver.findElement(By.xpath("(//div[@class='J13SB_42 cl_fff cpn_code_box bg_29'])["+j+"]"));
                    	   offerCode = code.getText();
                    	  System.out.println(offerCode);
                    	  keyValueMap.put(discountValue, offerCode);
                          System.out.println("Discount value :"+discountValue+" Discount Code :"+offerCode);
                    	
                      }
//                      else if (text.contains("Extra Rs.50 Off* on minimum purchase worth Rs.999")) {
//                    	  
//                    	  WebElement code = driver.findElement(By.xpath("(//div[@class='J13SB_42 cl_fff cpn_code_box bg_29'])["+k+"]"));
//                    	  continue;
//                      }//||text.contains("Flat")) { 
                      else  if((text.contains("FLAT")&& text.contains("FOR NON CLUB"))){
                   	//div[@class='cpncode_details_row']//div[@class='cpn_code_copy_block']//div
                    	  WebElement discountCode1 = driver.findElement(By.xpath("(//div[@class='J13SB_42 cl_fff cpn_code_box bg_29'])["+j+"]"));
                    	   discountCode = discountCode1.getText();
                    	  System.out.println(discountCode); 
                    	  if((text.contains("FLAT")&& text.contains("FOR NON CLUB"))) {
                    		  String Nonclub="For Non Club";
                    		  keyValueMap.put(Nonclub, discountCode);
                              System.out.println("Discount value :"+Nonclub+" Discount Code :"+discountCode);  
                    	  }
                    	  else  if(text.contains("Flat"))
                    	  {
                    		  WebElement discountCode3 = driver.findElement(By.xpath("(//div[@class='J13SB_42 cl_fff cpn_code_box bg_29'])["+j+"]"));
                       	      String discountCodeFlat = discountCode3.getText();
                       	      System.out.println(discountCodeFlat);
                    		  String FlatOff="Flat Off";
                    		  keyValueMap.put(FlatOff, discountCodeFlat);
                              System.out.println("Discount value :"+FlatOff+" Discount Code :"+discountCode);  

                      }
//                      keyValueMap.put(discountValue, offerCode);
//                      System.out.print("Discount value :"+discountValue+" Discount Code :"+offerCode);
                      }	
                      
                      }
                      //}
                      //for(int j=1;j<=divCount;j++) {
                    	// String[] OfferCouponValueTxtArray = new String[divCount];
                    	// WebElement OffervalueText= driver.findElement(By.xpath("(//span[@class='nn-club-tx'])["+j+"]"));    //(//div[@class='club-txt-logo']) for coupon value text
                    	// OfferCouponValueTxtArray[j] = OffervalueText.getText();
                    //	 System.out.println("Coupon Value "+ j +":"+OfferCouponValueTxtArray[j]);
                    	
                    	 
                    	// if(OfferCouponValueTxtArray[j].contains("All Users")||OfferCouponValueTxtArray[j].contains("FOR NON CLUB")||OfferCouponValueTxtArray[j].contains("ALL USERS")||OfferCouponValueTxtArray[j].contains("minimum purchase"));
                    	// {
//                    		 if(OfferCouponValue.contains("All Users")) {
//                    			 
//                    			 String[] parts = OfferCouponValue.split("All Users - ");
//
//                    		        if (parts.length >= 2) {
//                    		            // Extracting the discount portion
//                    		            String discountText = parts[1]; // This will be "Extra 5% Off* on Diapers"
//
//                    		            // Further splitting to get just the discount value
//                    		            String[] discountParts = discountText.split(" ");
//                    		            if (discountParts.length >= 2) {
//                    		                String discountValue = discountParts[1]; // This will be "5%"
//                    		                System.out.println("Discount value: " + discountValue);
//                    		            }
//                    		 }}
//                    		 if(OfferCouponValueTxtArray[j].contains("All Users")) {
//                    			 int startIndex = OfferCouponValueTxtArray[j].indexOf("5% Off");
//
//                    		        // Extract "5% Off" from the offer text
//                    		        String discountValue = OfferCouponValueTxtArray[j].substring(startIndex, startIndex + 6).trim(); // Adjust the length if "5% Off" changes
//
//                    		        // Print the extracted discount value
//                    		        System.out.println("Discount value: " + discountValue);}
//                    		 if(OfferCouponValueTxtArray[j].contains("minimum purchase")) {
//                    			 int startIndex = OfferCouponValueTxtArray[j].indexOf("minimum purchase");
//
//                    		        // Extract "Minimum purchase" from the offer text
//                    		        String minimumPurchaseText = OfferCouponValueTxtArray[j].substring(startIndex, startIndex + "minimum purchase".length()).trim(); 
//
//                    		        // Print the extracted "Minimum purchase" text
//                    		        System.out.println("Minimum purchase text: " + minimumPurchaseText);
//                    		    }
                    			 
//                    		
//                    		 WebElement OfferClick=driver.findElement(By.xpath("(//div[@class='J13SB_42 cl_fff cpn_code_box bg_29'])["+j+"]"));
//                    		 String OfferText=OfferClick.getText();
//                    		 System.out.println("Coupon Code: "+OfferText);
//                    	 }
//                    	 
//                      }
                    
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
                      //Add to cart
                      //for (int j = 0; j < 150; j++) {
                         //  driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                       	driver.findElement(By.xpath("//*[@id=\"p_breadcrumb\"]/div[2]/div/div[2]/div[6]/div/div[2]/div[2]/div[1]/div/span[1]")).click();
                   		
                       	 Thread.sleep(1000); 
                       	
                   		driver.findElement(By.xpath("//*[@id=\"p_breadcrumb\"]/div[2]/div/div[2]/div[6]/div/div[2]/div[2]/div[1]/div/span[2]")).click();
                   		
                          // driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                          // break;
                     //  }
                      }
                      catch(NoSuchElementException e){
                      	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                      	
                      	try {
                      		Thread.sleep(500);
                          	driver.findElement(By.xpath("/html/body/div[5]/div/div[2]/div[1]/div[1]/div[2]/div[2]/div[1]/div/span[1]")).click();
                       		
                          	 Thread.sleep(1000); 
                          	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                      		driver.findElement(By.xpath("/html/body/div[5]/div/div[2]/div[1]/div[1]/div[2]/div[2]/div[1]/div/span[2]")).click();
                      		
                      	}
                      	//Go to cart
                      	catch(NoSuchElementException ex) {
                      		Thread.sleep(2000);
                          	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                          	driver.findElement(By.xpath("(//span[@class='step2 M16_white'])[1]")).click();
                       		
                          	 Thread.sleep(1000); 
                          	// driver.findElement(By.xpath("//select[@id =\"product_qty\"]/ancestor::span")).click();
                      		driver.findElement(By.xpath("//*[@id=\"p_breadcrumb\"]/div[2]/div/div[2]/div[5]/div/div[2]/div[2]/div[1]/div/span[2]")).click();
                      		
                      	}
                  		
                      }
                      //catch the 2 product has the different xpath for add to card
                      
                      //Add to cart option not available 
                      catch(Exception ex){
                      	
                      	//count = 1;
                      	
                      	WebElement nameElement = driver.findElement(By.id("prod_name"));
                          newName = nameElement.getText();
                          
                          Thread.sleep(500);
                          
                          WebElement mrp = driver.findElement(By.xpath("//*[@id=\"original_mrp\"]"));
                          mrpValue = mrp.getText();
                          
                          Thread.sleep(500);
                          
                          WebElement sp = driver.findElement(By.xpath("//*[@id=\"prod_price\"]"));
                         spValue = sp.getText();
                         System.out.println("+++++++++++" + spValue);
                          finalSp = spValue;
                         
                         count=1;
                      }
                      System.out.println("==================="+count+"===================");
                      
                      //screenshot
                      BlinkitId screenshot = new BlinkitId();

                      try {
          				screenshot.screenshot(driver, "Firstcry", id);
          			} catch (Exception e) {
          				e.fillInStackTrace();
          			
          			}
                      
                      
                      finalSp = spValue;
                      
                      
                      System.out.println("==================="+finalSp+"===================");
                      
                    //  finalSp = Integer.parseInt(spValue);
                      //sp scrap
                      String rateValue = "";
                      
                    //  if(count != 1 ) {
                      Thread.sleep(2000);
                      WebElement rate = driver.findElement(By.id("NetPayment"));
                      rateValue = rate.getText();
                      System.out.println(rateValue);
                      finalSp = rateValue;
                     // }
                      
                      //coupon code copy
                      //rateValue = spValue;
                      Thread.sleep(5000);
                     WebElement coupon=driver.findElement(By.xpath("//*[@id=\"CouponCode\"]"));
                     coupon.click();
                     coupon.clear();
                     coupon.sendKeys(OfferCode);
                     
                     // coupon code apply
                  
                     Thread.sleep(1000);
                     WebElement aplcoupon=driver.findElement(By.xpath("	//*[@id=\"apply-coupon-button\"]"));
                     aplcoupon.click();
                     
                     Thread.sleep(2000);
                     
                     
                     //After apply coupon code sp value scrap
                     
                     String OFFrateValue="";
                     WebElement OFFrate = driver.findElement(By.id("NetPayment"));
                     OFFrateValue = OFFrate.getText();
                     System.out.println(OFFrateValue);
                     sp1 = OFFrateValue;
                      
                    boolean success = false;
                      int attempts = 0;
                      	//if(count != 1) {
                      while (!success && attempts < 3) { 
                          try {
                          	Thread.sleep(4000); 
                              //driver.findElement(By.xpath("//*[@id=\"garem_3312344\"]")).click();
                              
                           WebElement remove = driver.findElement(By.className("remove-icon"));
                           
                           remove.click();

                              success = true;
                          } catch (Exception e) {
                             
                              e.printStackTrace();
                              attempts++;
                              Thread.sleep(5000); 
                          }
                          
                      }  
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
                      

                      System.out.println("Failed to extract data for URL: " + url);
                      
                  }
              }
              
              // Write results to Excel file
              FileOutputStream outFile = new FileOutputStream(".\\Output\\Firstcry outputData.xlsx");
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