package Shopping;

import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;

public class bigBasketAll {

	public static void main(String[] args)throws Exception {


		System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
	        WebDriver driver = new ChromeDriver();
	        
	        String url="https://www.bigbasket.com";
	        driver.get(url);
	        driver.manage().window().maximize();
	       
	        WebElement category=driver.findElement(By.xpath("(//div[@class='relative h-full'])[4]"));
	        Thread.sleep(5000);
	       
	        category.click();
	        
	        Thread.sleep(1000);
	        
	        List<WebElement> ulElements = driver.findElements(By.xpath("/html/body/div[2]/div[1]/header[2]/div[2]/div[1]/div[1]/div/div/nav/ul/li"));
	       
	        int hi = ulElements.size();
	       System.out.println(hi);
	       
	       Workbook workbook = new XSSFWorkbook();
	    //   Sheet sheet = workbook.createSheet("Results");
	       org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet("Product Data");
 
	       
	       for (int i = 0; i <= 12; i++) {
	    	    WebElement ulElement = ulElements.get(i);
	    	    System.out.println("UL Element " + (i + 1) + ":");
	    	    ulElement.click();
	    	    Thread.sleep(10000);

	    	    List<WebElement> sepcates = driver.findElements(By.xpath("//li[@class='FilterByCategory___StyledLi-sc-c0pkxv-3 LRJmo']//a"));
	    	    for (int j = 0; j < sepcates.size(); j++) {
	    	        int retries = 3;
	    	        int w = 0;

	    	        while (retries > 0) {
	    	            try {
	    	                System.out.println(w);
	    	                System.out.println(sepcates.size());
	    	                WebElement sepcate = sepcates.get(w);
	    	                sepcate.click();
	    	                Thread.sleep(10000);
	    	                retries = 0; // Exit the loop if click was successful
	    	            } catch (StaleElementReferenceException e) {
	    	                // Retry clicking the element
	    	                retries--;
	    	                System.out.println("Retrying...");
	    	            }
	    	        }

	    	        List<WebElement> Category3 = driver.findElements(By.xpath("//li[@class='FilterByCategory___StyledLi-sc-c0pkxv-3 LRJmo']//a"));
	    	        for (int k = 0; k < Category3.size(); k++) {
	    	            int retries1 = 3; // Number of retries
	    	            int q = 0;

	    	            while (retries1 > 0 && q < Category3.size()) {
	    	                try {
	    	                    System.out.println(Category3.size());
	    	                    WebElement cate3 = Category3.get(q);
	    	                    Thread.sleep(10000);
	    	                    cate3.click();
	    	                    Thread.sleep(10000);
	    	                    q++; // Move to the next element
	    	                } catch (StaleElementReferenceException e) {
	    	                    // Retry clicking the element
	    	                    retries1--;
	    	                    System.out.println("Retrying...");
	    	                }
	    	            }
	    	            
	    	            // Scraping product details
	    	            Thread.sleep(5000);
	    	            List<WebElement> categoryLinks = driver.findElements(By.xpath(".//section[@class='z-10 ']//li"));
	    	            System.out.println("=============");
	    	            System.out.println(categoryLinks.size());
	    	            System.out.println("=============");

             /*         for (WebElement categoryLink : categoryLinks) {
                      	String productBrandName = "NA";
                          String productUrl = "NA";
                          String productNameResult = "NA";
                          String originalMrp = "NA";
                          String originalUom = "NA";
                          String spValue = "NA";
                          String offer = "NA";
                	  
                	  Thread.sleep(20000); // Add wait here if needed
                    
                	  for(int t = 0; t < 200; t++) {
                		  //products Link
                      try {
                          productUrl = categoryLink.findElement(By.xpath(".//div[@class='DeckImage___StyledDiv-sc-1mdvxwk-1 jbskZj']//a")).getAttribute("href");
                      } catch (Exception e) {
                      	 e.printStackTrace();
                      }
                	  }
                      Thread.sleep(1000);
                      //Product Brand
                      for (int a = 0; a < 200; a++) {
                      	 try {
                               productBrandName = categoryLink.findElement(By.xpath(".//h3//span[@class='Label-sc-15v1nk5-0 BrandName___StyledLabel2-sc-hssfrl-1 gJxZPQ keQNWn']")).getText();
                            //  System.out.println(productBrandName);
                               break;
                           } catch (Exception e) {
                          	 if (a==199) {
                          		 e.printStackTrace();
          					}
                           	
                           }
          			}
                     
                      for (int o = 0; o < 200; o++) {
                     // Thread.sleep(10000);
                      //Product name
                      try {
                          productNameResult = categoryLink.findElement(By.xpath(" .//h3[@class = 'block m-0 line-clamp-2 font-regular text-base leading-sm text-darkOnyx-800 pt-0.5 h-full']")).getText();
                       //   System.out.println(productNameResult);                      
                      break;
                      } catch (Exception e) {
                      	 e.printStackTrace();
                      }
                      
                      Thread.sleep(1000);
                      
                      }
                      //mrp
                      
                      for(int b=0;b<200;b++) {
                      try {
                          originalMrp = categoryLink.findElement(By.xpath(" .//span[@class = 'Label-sc-15v1nk5-0 Pricing___StyledLabel2-sc-pldi2d-2 gJxZPQ hsCgvu']")).getText().replace("₹", "");
                         // System.out.println(originalMrp);
                          break;
                      } catch (Exception e) {
                    	  originalMrp = "NA";
                      	if (b==199) {
                   		   e.printStackTrace();
          				}
                          
                      }
                      }

                      Thread.sleep(1000);
                      //sp
                      
                      for(int c=0;c<200;c++) {
                      	
                      
                      try {
                          spValue = categoryLink.findElement(By.xpath(" .//span[@class = 'Label-sc-15v1nk5-0 Pricing___StyledLabel-sc-pldi2d-1 gJxZPQ AypOi']")).getText().replace("₹", "");
                      //    System.out.println(spValue);
                          break;
                      } catch (Exception e) {
                      	if (c==199) {
                      		 e.printStackTrace();
          				}
                      	
                      }
                  }
                  
                     Thread.sleep(1000);
                     for (int f = 0; f < 200; f++) {
                  	 //uom
                         try {
                           //  originalUom = categoryLink.findElement(By.xpath(" .//span[@class = 'Label-sc-15v1nk5-0 gJxZPQ truncate']")).getText();
                        	 originalUom = categoryLink.findElement(By.xpath("//h3[contains(text(),'"+productNameResult+"')]/ancestor::a/following-sibling::div[2]//span")).getText();
                        	 //  System.out.println(originalUom);
                             break;
                         } catch (Exception e) {
                      	   if (f==199) {
                      		   e.printStackTrace();
          				}
                         	
                         }
          		}
                      
                      Thread.sleep(1000);
                    //uom
                      for (int g = 0; g < 200; g++) {
                      	try {
                              offer = categoryLink.findElement(By.xpath(".//h3[contains(text(),'"+productNameResult+"')]/ancestor::h3/parent::div/div[1]//div[contains(@class,'Offers')]//span[contains(text(),'% ')]")).getText().replace("OFF", "Off");
                            //  System.out.println(offer);
                              break;
                          } catch (Exception e) {
                          	 if (g==199) {
                        		   e.printStackTrace();
            				}
                          }
                	 
                  }
                      System.out.println("Product Brand: " + productBrandName);
                      System.out.println("Product URL: " + productUrl);
                      System.out.println("Product Name: " + productNameResult);
                      System.out.println("Original MRP: " + originalMrp);
                      System.out.println("SP Value: " + spValue);
                      System.out.println("Original UOM: " + originalUom);
                      System.out.println("Original Offer: " + offer);
                      
                      
                      
                   // Write data to Excel
                      Row row = sheet.createRow(sheet.getLastRowNum() + 1);

                      Cell brandCell = row.createCell(0);
                      brandCell.setCellValue(productBrandName);

                      Cell urlCell = row.createCell(1);
                      urlCell.setCellValue(productUrl);

                      Cell nameCell = row.createCell(2);
                      nameCell.setCellValue(productNameResult);

                      Cell mrpCell = row.createCell(3);
                      mrpCell.setCellValue(originalMrp);

                      Cell spCell = row.createCell(4);
                      spCell.setCellValue(spValue);

                      Cell uomCell = row.createCell(5);
                      uomCell.setCellValue(originalUom);

                      Cell offerCell = row.createCell(6);
                      offerCell.setCellValue(offer);
                      
                      }  */
                  
                  Thread.sleep(10000);
                  
               //   WebDriverWait wait = new WebDriverWait(driver, Seconds(10));
               //   wait.until(ExpectedConditions.elementToBeClickable(category));
                 
          /*       // Thread.sleep(10000);
                  for(int =0)
                  WebElement backClick=driver.findElement(By.xpath("/html/body/div[2]/div[1]/div[5]/div[2]/section[1]/div/div[1]/div[1]/a/svg/g/path"));
                  
                  backClick.click();   */
                  }
                  
                  
               //   
                  Thread.sleep(10000);
                  
                  WebElement backClick = driver.findElement(By.xpath("//*[local-name() = 'g' and contains(@mask, 'url(#arrow-left_svg__a)')]/*[local-name() = 'path']"));
                  
                  Thread.sleep(2000);
                  
                  backClick.click();
                  
      	          //  driver.navigate().back();
      	            
      	        //  Thread.sleep(10000);
      	            
                  }
                category.click(); 
          }
	        
	    // Save the workbook to a file
	       try {
           	// for store the multiple we can use the time to store the multiple files
               SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
               String timestamp = dateFormat.format(new Date());
               String outputFilePath = ".\\Output\\Bigbasket_Fully_OutputData_" + timestamp + ".xlsx";
               
               // Write results to Excel file
               FileOutputStream outFile = new FileOutputStream(outputFilePath);
               workbook.write(outFile);
               outFile.close();
               
               System.out.println("Output file saved: " + outputFilePath);
           } catch (Exception e) {
               e.printStackTrace();
           }
	        
	       finally {
	            if (driver != null) {
	            	System.out.println("DoNe DoNe Scraping DoNe");
	                driver.quit();
	            }
	    
	}
	

}
}

