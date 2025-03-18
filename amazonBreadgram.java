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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class amazonBreadgram {

    public static void main(String[] args) throws Exception{
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        int urlCount = 0; 
        WebElement ulElement = null;
         
        try {
            // Read URLs from Excel file
            String filePath = ".\\input-data\\amazonBreadgram.xlsx";
            FileInputStream file = new FileInputStream(filePath);
            Workbook urlsWorkbook = new XSSFWorkbook(file);
            Sheet urlsSheet = urlsWorkbook.getSheet("Sheet3");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

	            List<String> inputPid = new ArrayList<>(),
	            		uRL = new ArrayList<>();
	            
            // Extract URLs from Excel
            for (int i = 0; i < rowCount; i++) {
                Row row = urlsSheet.getRow(i);
                
                
                  if (i == 0) {
                continue;
            }    
                
                
                Cell inputPidCell = row.getCell(0);
                Cell urlCell = row.getCell(1);
            
                if (urlCell != null && urlCell.getCellType() == CellType.STRING) {
                    String url = urlCell.getStringCellValue();
                    String id = (inputPidCell != null && inputPidCell.getCellType() == CellType.STRING) ? inputPidCell.getStringCellValue() : "";
                
                    inputPid.add(id);
                    uRL.add(url);
                    
                    urlCount++;
                  
                }
            }
            // Create Excel workbook for storing results
            Workbook resultsWorkbook = new XSSFWorkbook();
            Sheet resultsSheet = resultsWorkbook.createSheet("Results");

            Row headerRow = resultsSheet.createRow(0);
            
            
            headerRow.createCell(0).setCellValue("InputPid");
            headerRow.createCell(1).setCellValue("URL");
            headerRow.createCell(2).setCellValue("Complete Name");
           
           
           
            
            int rowIndex = 1;

            int headercount = 0;
            
            for (int i = 0; i < uRL.size(); i++) {
                String id = inputPid.get(i);
                String url = uRL.get(i);
                
                
                
                try {
                	
                	  if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
                          // Set "NA" values in all three columns
                          Row resultRow = resultsSheet.createRow(rowIndex++);
                          resultRow.createCell(0).setCellValue(id);
                          resultRow.createCell(0).setCellValue(url);
                          
                          System.out.println("Skipped processing for URL: " + url);
                          continue; 
                      }
                	  
                    driver.get(url);
                    driver.manage().window().maximize();
                    Thread.sleep(7000);
                    try {
                    ulElement = driver.findElement(By.xpath("/html/body/div[1]/div/div[3]/div/div/ul")); // Adjust the locator as per your HTML structure
                    }
                    catch(Exception e) {
                    	try {
                    ulElement = driver.findElement(By.xpath("/html/body/div[1]/div/div[7]/div/div/ul"));
                    	}
                    	catch(Exception r) {
                    		try {
                    		ulElement = driver.findElement(By.xpath("/html/body/div[1]/div/div[5]/div/div/ul"));
                    		}
                    		catch(Exception y) {
                    			try {
                    		ulElement = driver.findElement(By.xpath("/html/body/div[1]/div/div[6]/div/div/ul"));
                    			}
                    			catch(Exception j) {
                    				ulElement = driver.findElement(By.xpath("/html/body/div[1]/div/div[8]/div/div/ul"));
                    				
                    				}
                    		}
                    		
                    	}
                    }
                   
                 List<WebElement> liElements = ulElement.findElements(By.xpath(".//li/span/a[@class='a-link-normal a-color-tertiary']")); // Use .//li to start the search from current node

                 System.out.println("Number of breadcrumb elements found: " + liElements.size());

                 StringBuilder breadcrumbs = new StringBuilder();
                 for (int j = 0; j < liElements.size(); j++) {
                     WebElement aElement = liElements.get(j); // Get the <a> element at index j
                     String breadcrumbText = aElement.getText().trim();
                     System.out.println("Text: " + breadcrumbText);

                     breadcrumbs.append(breadcrumbText);

                     if (j < liElements.size() - 1) {
                         breadcrumbs.append(" › ");
                         
                         
                     }
                 }

                 String Complete = breadcrumbs.toString();
                 
                 System.out.println("Complete Breadcrumbs: " + Complete);

                    

                    System.out.println("headercount = " + headercount);
                    
                    headercount++;
                    
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    
                    resultRow.createCell(0).setCellValue(id);
                    resultRow.createCell(1).setCellValue(url);
                    resultRow.createCell(2).setCellValue(Complete);
                    
                    System.out.println("Data extracted for URL: " + url);
                } catch (Exception e) {
                    e.printStackTrace();
                    
                    Row resultRow = resultsSheet.createRow(rowIndex++);
                    resultRow.createCell(0).setCellValue(id);
                   resultRow.createCell(1).setCellValue(url);
                    resultRow.createCell(2).setCellValue("NA"); 

                    System.out.println("Failed to extract data for URL: " + url);
                    
                }
            }
            try {
            	// for store the multiple we can use the time to store the multiple files
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Amazon_Breadgram_OutputData_FirstHalf" + timestamp + ".xlsx";
                
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