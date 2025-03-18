package Shopping;

import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;

import CommonUtility.BlinkitId;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class BBdemo {

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
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\BB28feb.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Feb29");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

	            List<String> inputPid = new ArrayList<>(),InputCity = new ArrayList<>(),InputName = new ArrayList<>(),InputSize = new ArrayList<>(),NewProductCode = new ArrayList<>(),
	            		uRL = new ArrayList<>(),UOM = new ArrayList<>(),Mulitiplier = new ArrayList<>(),Availability = new ArrayList<>(),Pincode = new ArrayList<>(),NameForCheck = new ArrayList<>();
	            
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
                Cell pinCodeCell = row.getCell(9);        
                Cell oldNameCell = row.getCell(10);    
                
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
                    String uom = (uomCell != null && uomCell.getCellType() == CellType.STRING) ? uomCell.getStringCellValue() : "";
                    String mulitiplier = (multiplierCell != null && multiplierCell.getCellType() == CellType.STRING) ? multiplierCell.getStringCellValue() : "";
                    String availability = (availabilityCell != null && availabilityCell.getCellType() == CellType.STRING) ? availabilityCell.getStringCellValue() : "";
                    String locationSet = (pinCodeCell != null && pinCodeCell.getCellType() == CellType.STRING) ? pinCodeCell.getStringCellValue() : "";
                    String namecheck = (oldNameCell != null && oldNameCell.getCellType() == CellType.STRING) ? oldNameCell.getStringCellValue() : "";
                    
                    inputPid.add(id);
                    InputCity.add(city);
                    InputName.add(name);
                    InputSize.add(size);
                    NewProductCode.add(productCode);
                    uRL.add(url);
                    UOM.add(uom);
                    Mulitiplier.add(mulitiplier);
                    Availability.add(availability);
                    Pincode.add(locationSet);
                    NameForCheck.add(namecheck);
                    
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
            headerRow.createCell(18).setCellValue("NameForCheck");
            
            int rowIndex = 1;
            int headercount = 0;
            String currentPin =null;
            
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
                String locationSet = Pincode.get(i);
                String namecheck = NameForCheck.get(i);
                
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
                          resultRow.createCell(12).setCellValue("NA");
                          
                          System.out.println("Skipped processing for URL: " + url);
                          continue; // Skip to the next iteration
                      }
                	  
                   
                 //  if (currentPin == null || !currentPin.equals(locationSet)) {
                	  
              		  //driver.get("https://www.bigbasket.com/");
                      //driver.manage().window().maximize();
                	  driver.get(url);
                      driver.manage().window().maximize();  
                      Thread.sleep(1000);
//                  //location sets 
//                WebElement location = driver.findElement(By.className("//div[@class='Header___StyledAddressDropdown2-sc-19kl9m3-2 eyaQtr']"));
//                Thread.sleep(2000);
//               location.click();					
//      //
//      				//*[@id="headlessui-menu-items-:r2c:"]/div[2]/input
//      			System.out.println(locationSet);//*[@id="headlessui-menu-button-:r0:"]//*[@id=\"headlessui-menu-button-:r9:\"]
                 //span[@class='Label-sc-15v1nk5-0 gJxZPQ w-full inline text-sm text-darkOnyx-600 leading-md truncate']
//      				String tempPinNumber = "";
//      				for (int j = 0; j < 150; j++) {       
//      					try {
//      					//driver.findElement(By.className("//span[@class='Label-sc-15v1nk5-0 gJxZPQ ml-1.5 flex-1 text-sm text-darkOnyx-900 truncate font-semibold leading-base']")).click();
//      						
//      					driver.findElement(By.xpath("//*[@id=\"headlessui-menu-items-:rn:\"]/div[2]/input")).click();
//      				
//      				//driver.findElement(By.xpath("//*[@id=\"headlessui-menu-items-:r14:\"]/div[2]/input")).clear();
//      				
//      				//System.out.println("print the crt pin number" + locationSet);
//      				
//      			//	String crtPin = locationSet;
//      				//driver.findElement(By.xpath("//*[@id=\"headlessui-menu-items-:r14:\"]/div[2]/input"))
//      						//.sendKeys("382424");
//      				Thread.sleep(4000);
//      			
//      				//driver.findElement(By.xpath("/html/body/div[3]/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div[2]/div[1]/div")).click();
//      			///	Thread.sleep(2000);
//      				
//      				driver.findElement(By.xpath("/html/body/div[3]/div[2]/div[2]/div/div/div/div/div/div[2]/div/div/div/div[2]/div/button[2]")).click();
//      				
//      				break;
//      					} catch (Exception e) {
//      						e.getCause();	
//      						if (j == 300) {
//      							Assert.fail(e.getMessage());
//      						}
//      					}
//      				}
                     
                    
                    WebElement location=  driver.findElement(By.xpath("//*[@id=\"siteLayout\"]/header[2]/div[1]/div[2]/div[1]/div[1]/div"));
        	        Thread.sleep(2000);
                    location.click();
                    Thread.sleep(2000);
                    
                  try {
                  WebElement pcode=
                 		driver.findElement(By.xpath("(//input[@placeholder='Search for area or street name'])[2]"));
                     Thread.sleep(5000);
                     pcode.click();
                     pcode.sendKeys(locationSet);
                  }catch (Exception e) {
                	  WebElement pcode=
                       		driver.findElement(By.xpath("(//input[@placeholder='Search for area or street name'])[1]"));
                           Thread.sleep(5000);
                           pcode.click();
                           pcode.sendKeys(locationSet);
                  }
                   try {
                     Thread.sleep(2000);
                     WebElement place=driver.findElement(By.xpath("/html/body/div[2]/div[1]/header[2]/div[1]/div[2]/div[1]/div[1]/div/div[1]/div[3]/ul/li[1]"));		//((//div)[33]/ul/li)[1]
                     Thread.sleep(2000);
                     place.click();
                     }catch (NoSuchElementException e) {
                    	 Thread.sleep(2000);
                         WebElement place=driver.findElement(By.xpath("/html/body/div[2]/div[1]/header[2]/div[1]/div[2]/div[1]/div[1]/div/div[1]/div[3]/ul/li[2]"));		//((//div)[33]/ul/li)[1]
                         Thread.sleep(2000);
                         place.click();
					}
                   
                  // }
                      
                  
                    try {
                    WebElement nameElement = driver.findElement(By.xpath("//*[@id=\"siteLayout\"]/div/div/section[1]/div[2]/section[1]/h1"));
                    newName = nameElement.getText();
                    System.out.println(newName);
                    }
                    
                    catch(NoSuchElementException e) {
                    	WebElement nameElement = driver.findElement(By.xpath("/html/body/div[2]/div[1]/div/div/section[1]/div[2]/section[1]/h1"));
                    	newName = nameElement.getText();
                        System.out.println(newName);
                    }
                    System.out.println("headercount = " + headercount);
                    
                    headercount++;
                    
                    try {
                    WebElement mrp = driver.findElement(By.xpath("/html/body/div[2]/div[1]/div/div/section[1]/div[2]/section[1]/table/tr[2]/td[1]"));
                    originalMrp1 = mrp.getText();
                    mrpValue = originalMrp1.replace("₹", "");
                    System.out.println(mrpValue);
                    }
                    
                    catch(NoSuchElementException e){ 
                    	try {
                    		
                    		WebElement mrp = driver.findElement(By.xpath("/html/body/div[2]/div[1]/div/div/section[1]/div[2]/section[1]/table/tr[2]/td[1]"));
                            originalMrp2 = mrp.getText();
                            mrpValue = originalMrp2.replace("₹", "");
                           // mrpValue = originalMrp2;                      
                            System.out.println(mrpValue);
                        
                    }
                    	catch(Exception ex) {
                    		try {
                    		WebElement mrp = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[2]/span[2]/span[2]"));
                           // WebElement mrp = driver.findElement(By.xpath("/html/body/div[2]/div/div[7]/div[3]/div[4]/div[12]/div/div/div[4]/div[2]/span/span[1]/span[2]/span/span[2]"));
                            originalMrp3 = mrp.getText();
                            mrpValue = originalMrp3.replace("₹", "");                      
                            System.out.println(mrpValue);
                    		}
                    		catch(Exception exx) {
                    			mrpValue = "NA";
                    		}
                    		}
                    	}
                    
                   try {
                    WebElement sp = driver.findElement(By.xpath("//*[@id=\"siteLayout\"]/div/div/section[1]/div[2]/section[1]/table/tr[2]/td[1]/text()[2]"));
                    originalSp1 = sp.getText();
                    spValue =  originalSp1.replace("₹", "");
                    System.out.println(spValue);
                   }
                   catch(Exception e) {
                	   
                	   try {
                	   WebElement sp = driver.findElement(By.xpath("//*[@id=\"siteLayout\"]/div/div/section[1]/div[2]/section[1]/table/tr[1]/td[1]/text()[2]"));
                       originalSp2 = sp.getText();
                       spValue =  originalSp2.replace("₹", "");
                       System.out.println(spValue);
                	   }
                       catch(Exception exx) {
                    	   spValue = "NA";
                       }
                   }
                   
                   // offer
                   try {
                       WebElement offer = driver.findElement(By.xpath("//*[@id=\"siteLayout\"]/div/div/section[1]/div[2]/section[1]/table/tr[3]/td[2]"));
                       offerValue = offer.getText();
                       
                     /*  Pattern pattern = Pattern.compile("\\((.*?)\\)");
                       Matcher matcher = pattern.matcher(offer1);
                       
						if(matcher.find()) { 
							  String offer2 = matcher.group(1);
							  offerValue = offer2.replace("%","% Off");
						  }   */
                       //offerValue = offer.getText();
                       System.out.println(offerValue);
                      }
                      catch(Exception e) {
                    	  try {
                   	   WebElement offer = driver.findElement(By.xpath("//*[@id=\"corePriceDisplay_desktop_feature_div\"]/div[1]/span[2]"));
                   	   String NewOffer = offer.getText();
                       offerValue = NewOffer.replace("-","").replace("%","% Off");
                          System.out.println(offerValue);
                    	  }
                    	  catch(Exception ex){
                    		  offerValue = "NA";
                    	  }
                      }    
                   
                   
                   		//Screenshots 
                      BlinkitId screenshot = new BlinkitId();
	                   try {
	       				screenshot.screenshot(driver, "Bigbasket", id);
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
                    resultRow.createCell(11).setCellValue(availability);
                    resultRow.createCell(12).setCellValue(" ");
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(offerValue);
                    resultRow.createCell(18).setCellValue(namecheck);
                    
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
                    resultRow.createCell(11).setCellValue(availability);
                    resultRow.createCell(12).setCellValue(" ");
                    resultRow.createCell(13).setCellValue(" ");
                    resultRow.createCell(14).setCellValue(" ");
                    resultRow.createCell(15).setCellValue(" ");
                    resultRow.createCell(16).setCellValue(" ");
                    resultRow.createCell(17).setCellValue(offerValue);
                    resultRow.createCell(18).setCellValue(namecheck);
                    

                    System.out.println("Failed to extract data for URL: " + url);
                    
                }
            }
            
            // Write results to Excel file
            FileOutputStream outFile = new FileOutputStream(".\\Output\\Bigbasketmar1.xlsx");
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





			
             
