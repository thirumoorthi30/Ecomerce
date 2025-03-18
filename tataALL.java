package Shopping;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

//import CommonUtility.BlinkitId;

public class tataALL {
    public static void main(String[] args) throws Exception {
    	 // Ensure this path points to the chromedriver.exe, not chrome.exe
        // System.setProperty("webdriver.chrome.driver", "./Drivers/chromedriver.exe");

         ChromeOptions options = new ChromeOptions();
//         options.addArguments("--start-maximized");
//         options.addArguments("--window-size=375,812");
//         options.addArguments("--user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 13_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0 Mobile/15E148 Safari/604.1");

         WebDriver driver = new ChromeDriver(options);
         driver.manage().window().maximize();
         WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

         List<String> InputName = new ArrayList<>(),
                  uRL = new ArrayList<>();

         Workbook resultsWorkbook = new XSSFWorkbook();
         Sheet resultsSheet = resultsWorkbook.createSheet("Results");
         createHeaderRow(resultsSheet);
 //This product is currently unavailable in your selected pincode
         int rowIndex = 1;
         String currentPin = null;

         try (FileInputStream file = new FileInputStream(".\\input-data\\T1mg.xlsx");
              Workbook urlsWorkbook = new XSSFWorkbook(file)) {

             Sheet urlsSheet = urlsWorkbook.getSheet("Sheet1");
             int rowCount = urlsSheet.getPhysicalNumberOfRows();

             // Extract URLs from Excel
             for (int i = 1; i < rowCount; i++) {
                 Row row = urlsSheet.getRow(i);
                 if (row.getCell(1) != null) {
                     String url;
                     if (row.getCell(1).getCellType() == CellType.STRING) {
                         url = row.getCell(1).getStringCellValue();
                     } else if (row.getCell(1).getCellType() == CellType.NUMERIC) {
                         url = String.valueOf(row.getCell(1).getNumericCellValue());
                     } else {
                         url = "NA"; // Handle cases where the cell is neither numeric nor string
                     }

                     
                     InputName.add(row.getCell(0).getStringCellValue());
                     uRL.add(url);
                    
                 }
             }

//             Set<String> usedPincodes = new HashSet<>();

             int ProductCOUNT = 0;

             // Main data extraction logic
             for (int i = 0; i < uRL.size(); i++) {
                 String url = uRL.get(i);
               //  String locationSet = Pincode.get(i); // Get the current pincode from the list
                 if (url == null || url.isEmpty() || url.equalsIgnoreCase("NA")) {
                	    writeResults(resultsSheet, rowIndex++, InputName.get(i),
                	            url, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA","NA");
                	    System.out.println("Skipped processing for URL: " + url);
                	    continue;
                	} 


                 try {
                     driver.get(url);
                     Thread.sleep(2000);
                     
               /*    //  Thread.sleep(200000);
                     if (currentPin == null || !currentPin.equals(locationSet)) {
                     	Thread.sleep(2000);
                     	driver.findElement(By.xpath("//div[@class='header_pincode__KryhE']")).click();
                     	Thread.sleep(2000);
                     	driver.findElement(By.xpath("//input[@id='pincodeInput']")).sendKeys(locationSet);
                     	Thread.sleep(3000);
                     	driver.findElement(By.xpath("(//div[@class='pincode-widget_pincode-right__TwcOu'])[1]")).click();
                     	Thread.sleep(2000);
                     	driver.findElement(By.xpath("//button[.='CONFIRM LOCATION']")).click();
                     	Thread.sleep(2000);
      		
                     	currentPin = locationSet;
                     	
                     	driver.get(url);  
                          }  */
                     Thread.sleep(3000);


                     System.out.println("PRODUCTCOUNT = " + ProductCOUNT);

                     
                     String newName = " ";
                     try {
                     newName = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1"))).getText();
                     System.out.println("newName ="+newName);
                     }
                     catch (Exception e) {
                     	newName = "NA";
 					}
                     
                     Thread.sleep(3000);
                     
                     // Extract MRP
                     String mrpValue = extractMRP(driver);
                     System.out.println("mrpValue = "+mrpValue);
                     // Extract SP
                     String spValue = extractSP(driver);
                     System.out.println("spValue = "+spValue);
                     // Extract offer
                     String offerValue = extractOffer(driver);
                     System.out.println("offerValue = "+offerValue); 
                     
                     String stripSize = extractStripSize(driver);
                     System.out.println("stripSize = "+stripSize);
                     
                     String delTime = extractDelTime(driver);
                     System.out.println("delTime = "+delTime);
                     
                     String marketName = extractMarketer(driver);
                     System.out.println("marketName = "+marketName);
                     
                     String saltName = extractSalt(driver);
                     System.out.println("saltName =" +saltName);
                     
                     String storageName = extractStorage(driver);
                     System.out.println("storageName = "+storageName);
                     
                     String proInfo = extractProductInfo(driver);
                     System.out.println("proInfo = "+proInfo);
                     
                     String useInfo = extractUsesInfo(driver);
                     System.out.println("useInfo = "+useInfo);
                     
                     String benefits = extractBenefits(driver);
                     System.out.println("benefits = "+benefits);
                     
                     String sideEffects = extractSideEffects(driver);
                     System.out.println("sideEffects = "+sideEffects);
                     
                     String prescription=extractprescription(driver);
                     System.out.println("Prescription = "+prescription);


        /*             int result;
                     try {
                         // Define the texts to check for
                         String[] textsToCheck = {
                                 "Out Of Stock",
                                 "This product is currently unavailable in your selected pincode"
                                 
                         };

                         // Get the page source
                         String pageSource = driver.getPageSource();

                         boolean isTextPresent = false;

                         // Check for the presence of any of the texts
                         for (String text : textsToCheck) {
                             if (pageSource.contains(text)) {
                                 isTextPresent = true;
                                 break;
                             }
                         }
                      


                         // Determine the result based on the presence of the text
                         result = isTextPresent ? 0 : 1;
                         System.out.println(result);
                     } catch (Exception e) {
                         System.out.println("Error checking availability: " + e.getMessage());
                         result = -1;
                     }
//                     BlinkitId screenshot = new BlinkitId();
//                 	  try {
//             				screenshot.screenshot(driver, "Dmart", productId);
//             			} catch (Exception e) {
//             				e.fillInStackTrace();
//             			
//             			}
                     String availability = String.valueOf(result);   */

                     ProductCOUNT++;

                     // Write results to the results sheet
                     writeResults(resultsSheet, rowIndex++, InputName.get(i), 
                             url, newName, mrpValue, spValue, stripSize, delTime, 
                             marketName, saltName, storageName, proInfo, useInfo, 
                             benefits, sideEffects, offerValue,prescription);
 

                     System.out.println("Data extracted for URL: " + url);
                 } catch (Exception e) {
                     e.printStackTrace();
                     System.out.println("Failed to extract data for URL: " + url);
                 }
             }
         } catch (Exception e) {
             e.printStackTrace();
             System.out.println("An error occurred during the extraction process.");
         } finally {
             // Write results to Excel file
             saveResultsToExcel(resultsWorkbook);

             // Clean up driver resources
             if (driver != null) {
                 System.out.println("Closing the driver.");
                 driver.quit();
             }
         }
     }

     private static void createHeaderRow(Sheet sheet) {
         Row headerRow = sheet.createRow(0);
         String[] headers = {"INPUT NAME", "URL", "PRODUCT NAME", "MRP", "SP", "STRIPE SIZE", "DELIVERY TIME", "MARKETER NAME",
                 "SALT", "STORAGE", "PRODUCT INFO", "USES", "BENEFITS", "SIDE EFFECTS", "OFFER","PRESCRIPTION"
                 };
         for (int i = 0; i < headers.length; i++) {
             headerRow.createCell(i).setCellValue(headers[i]);
         }
     }

     public static void writeResults(Sheet resultsSheet, int rowIndex, String inputName, String url, String newName, 
             String mrpValue, String spValue, String stripSize, String delTime, String marketName, 
             String saltName, String storageName, String proInfo, String useInfo, String benefits, 
             String sideEffects, String offerValue, String prescription ) {
Row row = resultsSheet.createRow(rowIndex);
row.createCell(0).setCellValue(inputName);
row.createCell(1).setCellValue(url);
row.createCell(2).setCellValue(newName);
row.createCell(3).setCellValue(mrpValue);
row.createCell(4).setCellValue(spValue);
row.createCell(5).setCellValue(stripSize);
row.createCell(6).setCellValue(delTime);
row.createCell(7).setCellValue(marketName);
row.createCell(8).setCellValue(saltName);
row.createCell(9).setCellValue(storageName);
row.createCell(10).setCellValue(proInfo);
row.createCell(11).setCellValue(useInfo);
row.createCell(12).setCellValue(benefits);
row.createCell(13).setCellValue(sideEffects);
row.createCell(14).setCellValue(offerValue);
row.createCell(15).setCellValue(prescription);

}


     private static void saveResultsToExcel(Workbook resultsWorkbook) {
         try {
             SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
             String timestamp = dateFormat.format(new Date());
             String outputFilePath = ".\\Output\\TATA_OutputData_sheet1" + timestamp + ".xlsx";

             // Write results to Excel file
             try (FileOutputStream outFile = new FileOutputStream(outputFilePath)) {
                 resultsWorkbook.write(outFile);
             }

             System.out.println("Output file saved: " + outputFilePath);
         } catch (Exception e) {
             e.printStackTrace();
             System.out.println("Failed to save the output file.");
         }
     }
     
     

     private static String extractMRP(WebDriver driver) {
         try {
             List<WebElement> elements = driver.findElements(By.xpath("//span[@class='PriceBoxPlanOption__margin-right-4___2aqFt PriceBoxPlanOption__stike___pDQVN']"));  //add to card

             if (elements.isEmpty()) {
            	 try {
                 elements = driver.findElements(By.xpath("//div[.='MRP']/following-sibling::div[1]"));  //sold out
            	 }
            	 catch (Exception e) {
            		 elements = driver.findElements(By.xpath("//span[.='MRP']/following-sibling::span[1]")); //add to csrd
				}
             }

             if (!elements.isEmpty()) {
                 String text = elements.get(0).getText();
                 return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
             } else {
                 return "NA";
             }
         } catch (Exception e) {
             return "NA";
         }
     }

     

     private static String extractSP(WebDriver driver) {
         try {
             List<WebElement> elements = driver.findElements(By.xpath("(//span[@class='PriceBoxPlanOption__offer-price___3v9x8 PriceBoxPlanOption__offer-price-cp___2QPU_'])[1]"));

             if (elements.isEmpty()) {
            	 try {
                 elements = driver.findElements(By.xpath("//div[.='MRP']/following-sibling::div[1]"));
            	 }
            	 catch (Exception e) {
            		 elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__price___dj2lv']//div[2]"));
				}
             }

             if (!elements.isEmpty()) {
                 String text = elements.get(0).getText();
                 return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
             } else {
                 return "NA";
             }
         } catch (Exception e) {
             return "NA";
         }
     }
     private static String extractOffer(WebDriver driver) {
    	    try {
                List<WebElement> elements = driver.findElements(By.xpath("//span[@class='PriceBoxPlanOption__discount___iN_jm']"));

                if (elements.isEmpty()) {
               	
                    elements = driver.findElements(By.xpath("//span[.='MRP']/following-sibling::span[2]"));
               	 
                }

                if (!elements.isEmpty()) {
                    String text = elements.get(0).getText();
                    return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
                } else {
                    return "NA";
                }
            } catch (Exception e) {
                return "NA";
            }
      
     }
     
     private static String extractStripSize(WebDriver driver) {
 	    try {
             List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugPriceBox__qty-wrapper___1RBzv']//div[2])[2]"));

             if (elements.isEmpty()) {
            	
                 elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            	 
             }

             if (!elements.isEmpty()) {
                 String text = elements.get(0).getText();
                 return isValidValue(text) ? text.trim() : "NA";
             } else {
                 return "NA";
             }
         } catch (Exception e) {
             return "NA";
         }
   
  }
     
     private static String extractDelTime(WebDriver driver) {
    	    try {
    	        List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='style__box___1ez55']//div)[1]"));

    	        if (elements.isEmpty()) {
    	            int result = -1;
    	            try {
    	                // Define the texts to check for
    	                String[] textsToCheck = {
    	                    "SOLD OUT",
    	                    "DISCONTINUED"
    	                };

    	                // Get the page source
    	                String pageSource = driver.getPageSource();

    	                boolean isTextPresent = false;

    	                // Check for the presence of any of the texts
    	                for (String text : textsToCheck) {
    	                    if (pageSource.contains(text)) {
    	                        isTextPresent = true;
    	                        break;
    	                    }
    	                }

    	                // Determine the result based on the presence of the text
    	                result = isTextPresent ? 0 : 1;
    	                if (isTextPresent) {
    	                    return "SOLD OUT or DISCONTINUED";
    	                }

    	                System.out.println(result);
    	            } catch (Exception e) {
    	                System.out.println("Error checking availability: " + e.getMessage());
    	                result = -1;
    	            }
    	            return "NA";
    	        } else {
    	            // If the element is found, proceed with the original logic
    	            String text = elements.get(0).getText();
    	            return isValidValue(text) ? text.trim() : "NA";
    	        }
    	    } catch (Exception e) {
    	        return "NA";
    	    }
    	}

     
     private static String extractMarketer(WebDriver driver) {
  	    try {
              List<WebElement> elements = driver.findElements(By.xpath("//div[.='Marketer']/following-sibling::div"));

              if (elements.isEmpty()) {
             	
                  elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
             	 
              }

              if (!elements.isEmpty()) {
                  String text = elements.get(0).getText();
                  return isValidValue(text) ? text.trim() : "NA";
              } else {
                  return "NA";
              }
          } catch (Exception e) {
              return "NA";
          }
    
   }
     private static String extractSalt(WebDriver driver) {
   	    try {
               List<WebElement> elements = driver.findElements(By.xpath("//div[.='SALT COMPOSITION']/following-sibling::div"));

               if (elements.isEmpty()) {
              	
                   elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
              	 
               }

               if (!elements.isEmpty()) {
                   String text = elements.get(0).getText();
                   return isValidValue(text) ? text.trim() : "NA";
               } else {
                   return "NA";
               }
           } catch (Exception e) {
               return "NA";
           }
     
    }
     
     private static String extractStorage(WebDriver driver) {
    	    try {
                List<WebElement> elements = driver.findElements(By.xpath("//div[.='Storage']/following-sibling::div"));

                if (elements.isEmpty()) {
               	
                    elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
               	 
                }

                if (!elements.isEmpty()) {
                    String text = elements.get(0).getText();
                    return isValidValue(text) ? text.trim() : "NA";
                } else {
                    return "NA";
                }
            } catch (Exception e) {
                return "NA";
            }
      
     }
     
     private static String extractProductInfo(WebDriver driver) {
 	    try {
             List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugOverview__content___22ZBX'])[1]"));

             if (elements.isEmpty()) {
            	
                 elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
            	 
             }

             if (!elements.isEmpty()) {
                 String text = elements.get(0).getText();
                 return isValidValue(text) ? text.trim() : "NA";
             } else {
                 return "NA";
             }
         } catch (Exception e) {
             return "NA";
         }
   
  }
     
     private static String extractUsesInfo(WebDriver driver) {
  	    try {
              List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugOverview__content___22ZBX']//ul)[1]"));

              if (elements.isEmpty()) {
             	
                  elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
             	 
              }

              if (!elements.isEmpty()) {
                  String text = elements.get(0).getText();
                  return isValidValue(text) ? text.trim() : "NA";
              } else {
                  return "NA";
              }
          } catch (Exception e) {
              return "NA";
          }
    
   }
     
     private static String extractBenefits(WebDriver driver) {
   	    try {
               List<WebElement> elements = driver.findElements(By.xpath("(//div[@class='DrugOverview__content___22ZBX']//div[@class='ShowMoreArray__tile___2mFZk'])[1]"));

               if (elements.isEmpty()) {
              	
                   elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
              	 
               }

               if (!elements.isEmpty()) {
                   String text = elements.get(0).getText();
                   return isValidValue(text) ? text.trim() : "NA";
               } else {
                   return "NA";
               }
           } catch (Exception e) {
               return "NA";
           }
     
    }
     
     private static String extractSideEffects(WebDriver driver) {
    	    try {
                List<WebElement> elements = driver.findElements(By.xpath("//div[@id='side_effects']//div[@class='DrugOverview__list-container___2eAr6 DrugOverview__content___22ZBX']/ul"));

                if (elements.isEmpty()) {
               	
                    elements = driver.findElements(By.xpath("//div[@class='DrugPriceBox__quantity___2LGBX']"));
               	 
                }

                if (!elements.isEmpty()) {
                    String text = elements.get(0).getText();
                    return isValidValue(text) ? text.trim() : "NA";
                } else {
                    return "NA";
                }
            } catch (Exception e) {
                return "NA";
            }
      
     }
     
     private static String extractprescription(WebDriver driver) {
 	    try {
             List<WebElement> elements = driver.findElements(By.xpath("//span[text()='Prescription Required']"));
             
             if(!elements.isEmpty()) {
                 String text = elements.get(0).getText();
                 return isValidValue(text) ? text.trim() : "NA";
             }else {
            	  return "NA";
             }}
             catch (Exception e) {
                 return "NA";
             }
        
  }

     
     
     private static boolean isValidValue(String value) {
         return value != null && !value.isEmpty() && !value.equals("₹");
     }
    

}
