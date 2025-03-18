package Shopping;

import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;

import CommonUtility.BlinkitId;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class myntraNew {

    /**
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception{
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        int count = 0;
        // int finalSp;
          String spValue = "";
          String finalSp = "";
          String offerValue = "NA";
          String newName = null;
          String mrpValue = null;
          String originalMrp1 = " ";
          String originalMrp2 = " ";
          String originalMrp3 = " ";
          String originalSp1 = " ";
          String originalSp2 = " ";
          String NewAvailability1 = " ";
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\nykaa input new.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Sheet1");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

	            List<String> inputPid = new ArrayList<>(),InputCity = new ArrayList<>(),InputName = new ArrayList<>(),InputSize = new ArrayList<>(),NewProductCode = new ArrayList<>(),
	            		uRL = new ArrayList<>(),UOM = new ArrayList<>(),Mulitiplier = new ArrayList<>(),Availability = new ArrayList<>();
	            
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
                Cell uomCell = row.getCell(6);
                Cell multiplierCell = row.getCell(7);
                Cell availabilityCell = row.getCell(8);
                
            //   Cell urlCell = row.getCell(0);
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
                    String uom = (uomCell != null && uomCell.getCellType() == CellType.STRING) ? uomCell.getStringCellValue() : "";
                    String mulitiplier = (multiplierCell != null && multiplierCell.getCellType() == CellType.STRING) ? multiplierCell.getStringCellValue() : "";
                    String availability = (availabilityCell != null && availabilityCell.getCellType() == CellType.STRING) ? availabilityCell.getStringCellValue() : "";
                    
                    inputPid.add(id);
                    InputCity.add(city);
                    InputName.add(name);
                    InputSize.add(size);
                    NewProductCode.add(productCode);
                    uRL.add(url);
                    UOM.add(uom);
                    Mulitiplier.add(mulitiplier);
                    Availability.add(availability);
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
            headerRow.createCell(9).setCellValue("UOM");
            headerRow.createCell(10).setCellValue("Multiplier");
            headerRow.createCell(11).setCellValue("Availability");
            headerRow.createCell(12).setCellValue("Commands");
            headerRow.createCell(13).setCellValue("Remarks");
            headerRow.createCell(14).setCellValue("Correctness");
            headerRow.createCell(15).setCellValue("Percentage");
            headerRow.createCell(16).setCellValue("Name");
            headerRow.createCell(17).setCellValue("Offer");
            
            int rowIndex = 1;

            int headercount = 0;
            
            for (int i = 0; i < uRL.size(); i++) {
                String id = inputPid.get(i);
                String city = InputCity.get(i);
                String name = InputName.get(i);
                String size = InputSize.get(i);
                String productCode = NewProductCode.get(i);
                String url = uRL.get(i);
                String uom = UOM.get(i);
                String mulitiplier = Mulitiplier.get(i);
                String availability = Availability.get(i);
                
                
                
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
                          resultRow.createCell(9).setCellValue("NA");
                          resultRow.createCell(10).setCellValue("NA");
                          resultRow.createCell(11).setCellValue("NA");
                          System.out.println("Skipped processing for URL: " + url);
                          continue; // Skip to the next iteration
                      }
                	  
                	  driver.manage().deleteAllCookies();
                	  
                	  
                    driver.get(url);
                    driver.manage().window().maximize();
                    //Thread.sleep(7000);
                    if(i == 0) {
                    	Thread.sleep(9000);
                    }
                    
					if(i == 0 ) {
						WebElement location = driver.findElement(By.className("css-i2hz85"));
						location.click();    
						String tempPinNumber = "";
						for (int j = 0; j < 50; j++) {
							try {
								
								
								//System.out.println("place 1 done ");
								
								
								driver.findElement(	
										By.className("css-i2hz85"))
										.sendKeys("110001");
								
								
								System.out.println("place pin entered ");
								
								
								
								
							/*	driver.findElement(
										By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[1]/div[2]/div/div[1]/div[6]/div[2]/div/div[2]/span")).clear();
								
								
								System.out.println("palce 3 done");
								
								
								//System.out.println("print the crt pin number" + InputPin);
								
								//String crtPin = InputPin;
								driver.findElement(
										By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[1]/div[2]/div/div[1]/div[6]/div[2]/div/div[2]/span"))
										.sendKeys("110001");
								
							/*	driver.findElement(
										By.xpath("//div[@id='GLUXZipInputSection']//input[@id='GLUXZipUpdateInput']"))
										.sendKeys(Keys.ENTER);   */
								
						/*		for (int k = 0; k <= 50; k++) {
									try {
										tempPinNumber = driver.findElement(By.xpath(
												"//*[@id=\"app\"]/div[1]/div[2]/div[1]/div[2]/div/div[1]/div[6]/div[2]/div/div[2]/span"))
												.getAttribute("value");
										if (tempPinNumber.equals("110001")) {
											break;
										}
									} catch (Exception e) {
										if (i == 50) {
											Assert.fail(e.getMessage());
										}
									}
								}   */
								driver.findElement(By.className("css-1ivth7r")).click();
								break;
								
							} catch (Exception e) {
								e.getCause();
									if (j == 300) {	
									Assert.fail(e.getMessage());
								}   
							}
						}
					}   
                    
                    try {
                    	
                    WebElement nameElement = driver.findElement(By.className("css-1gc4x7i"));
                    newName = nameElement.getText();
                    System.out.println(newName);
                    }
                    
                    catch(NoSuchElementException e) {
                    	
                    	WebElement nameElement = driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[1]/div[2]/div/div[1]/h1"));
                    	newName = nameElement.getText();
                        System.out.println(newName);
                    	
                    }
                    System.out.println("headercount = " + headercount);
                    
                    headercount++;
                    
                    try {
                    WebElement mrp = driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[1]/div[2]/div/div[1]/div[2]/div/span[1]/span"));
                    originalMrp1 = mrp.getText();
                    mrpValue = originalMrp1.replace("₹", "");
                    System.out.println(mrpValue);
                    
                    } 
                    
                    catch(NoSuchElementException e){ 
                    	try {
                    		WebElement mrp = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[2]/span/span[1]/span[2]/span/span[2]"));
                    	//	WebElement mrp = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[2]/span[2]/span[2]"));
                            originalMrp2 = mrp.getText();
                            mrpValue = originalMrp2.replace("₹", "");
                            System.out.println(mrpValue);
                        
                    }
                    	catch(Exception ex) {
                    		try {
                    			WebElement mrp = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[2]/span[2]/span[2]"));
                         //   WebElement mrp = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[2]/span/span[1]/span[2]/span/span[2]"));
                    		originalMrp3 = mrp.getText();
                    		if(originalMrp3.contains("₹")){   
                    			 mrpValue = originalMrp3.replace("₹", "");
                    		}else {
                    			mrpValue = originalMrp3;
                    		}   
                            System.out.println(mrpValue);
                    		}
                    		catch(Exception exx) {
                    			try {
                            		WebElement mrp = driver.findElement(By.xpath("/html/body/div[2]/div/div[7]/div[3]/div[4]/div[12]/div/div/div[4]/div[2]/span/span[1]/span[2]/span/span[2]"));
                            		originalMrp3 = mrp.getText();
                            		if(originalMrp3.contains("₹")){
                            			 mrpValue = originalMrp3.replace("₹", "");
                            		}else {
                            			mrpValue = originalMrp3;
                            		}   
                                    System.out.println(mrpValue);
                            		}
                            		catch(Exception exxR) {
                            			mrpValue = "NA";
                            		}
                    		}
                    		}
                    	}
                    
                   try {
                    WebElement sp = driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[1]/div[2]/div/div[1]/div[2]/div/span[2]"));
                    originalSp1 = sp.getText();
                    spValue =  originalSp1.replace("₹", "");
                    System.out.println(spValue);
                    
                   }
                   catch(Exception e) {
                	   try {
                	   WebElement sp = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[3]/span[2]/span[2]"));
                       originalSp2 = sp.getText();
                       spValue =  originalSp2.replace("₹", "");
                       System.out.println(spValue);
                	   }
                       catch(Exception exx) {
                    	   try {
                    		   WebElement sp = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[2]/span[2]/span[2]"));
                               originalSp2 = sp.getText();
                               spValue =  originalSp2.replace("₹", "");
                               System.out.println(spValue);
                    	   }
                    	   
                    		   catch(Exception ex) {
                    			   spValue = mrpValue;
                    		   }
                           }
                    	  
                       }
                   
                   //out of Stocks 
                   if(url.contains("NA")){
						String result = "NA";
					}	
					
					int result = 1;
					try {
					String xpathForNykaa = "//*[@id=\"app\"]/div[1]/div[2]/div[1]/div[2]/div/div[1]/div[6]/div/div/button";

					
					boolean isElementPresent = !driver.findElements(By.xpath(xpathForNykaa)).isEmpty();

			        result = isElementPresent ? 0 : 1;
			        
			        System.out.println(result);
					}
					catch(Exception e) {
						System.out.println(e.getMessage());
					}
					
					//int stock = result;
					NewAvailability1 = String.valueOf(result);
                   
                   
                   try {
                       WebElement offer = driver.findElement(By.xpath("//*[@id=\"app\"]/div[1]/div[2]/div[1]/div[2]/div/div[1]/div[2]/div/span[3]"));
                       offerValue = offer.getText();
                       System.out.println(offerValue);
                      }
                      catch(Exception e) {
                   	   WebElement offer = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[3]/span[2]/span[2]"));
                       offerValue = offer.getText();
                          System.out.println(offerValue);
                      }     
                   		//Screenshots 
                      BlinkitId screenshot = new BlinkitId();
	                   try {
	       				screenshot.screenshot(driver, "Nykaa", id);
	       			} catch (Exception e) {
	       				e.fillInStackTrace();
	       			
	       			}
                   
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    
                    resultRow.createCell(0).setCellValue(id);
                    resultRow.createCell(1).setCellValue(city);
                    resultRow.createCell(2).setCellValue(name); 
                    resultRow.createCell(3).setCellValue(size);
                    resultRow.createCell(4).setCellValue(productCode);
                    resultRow.createCell(5).setCellValue(url);
                    resultRow.createCell(6).setCellValue(newName);
                    resultRow.createCell(7).setCellValue(mrpValue);
                    resultRow.createCell(8).setCellValue(spValue);
                    resultRow.createCell(9).setCellValue(uom);
                    resultRow.createCell(10).setCellValue(mulitiplier);
                    resultRow.createCell(11).setCellValue(NewAvailability1);
                    resultRow.createCell(12).setCellValue(offerValue);
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(" ");
                    resultRow.createCell(18).setCellValue("NA");
                    
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
                    resultRow.createCell(9).setCellValue(uom);
                    resultRow.createCell(10).setCellValue(mulitiplier);
                    resultRow.createCell(11).setCellValue(NewAvailability1);
                    resultRow.createCell(12).setCellValue(offerValue);
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(" ");
                    resultRow.createCell(18).setCellValue("NA");

                    System.out.println("Failed to extract data for URL: " + url);
                    
                }
            }
            try {
            	// for store the multiple we can use the time to store the multiple files
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Nykaa_OutputData_ " + timestamp + ".xlsx";
                
                // Write results to Excel file
                FileOutputStream outFile = new FileOutputStream(outputFilePath);
                resultsWorkbook.write(outFile);
                outFile.close();
                
                System.out.println("Output file saved: " + outputFilePath);
            } catch (Exception e) {
                e.printStackTrace();
            }
           
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (driver != null) {
            	System.out.println("DoNe DoNe Scraping DoNe");
                driver.quit();
            }
        }
    }
}