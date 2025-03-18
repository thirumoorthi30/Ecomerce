package Shopping;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class bbNOW {
    public static void main(String[] args) throws Exception {
        // Ensure this path points to the chromedriver.exe, not chrome.exe
        System.setProperty("webdriver.chrome.driver", "./Drivers/chromedriver.exe");

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");
        options.addArguments("--window-size=375,812"); 
        options.addArguments("--user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 13_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0 Mobile/15E148 Safari/604.1");
       // options.addArguments("--auto-open-devtools-for-tabs");

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        
        int ProductCOUNT = 0;

        List<String> inputPid = new ArrayList<>(), InputCity = new ArrayList<>(), InputName = new ArrayList<>(),
                InputSize = new ArrayList<>(), NewProductCode = new ArrayList<>(), uRL = new ArrayList<>(),
                UOM = new ArrayList<>(), Mulitiplier = new ArrayList<>(), 
                Pincode = new ArrayList<>(), NameForCheck = new ArrayList<>();
        
        

        try (FileInputStream file = new FileInputStream(".\\input-data\\BB28feb.xlsx");
             Workbook urlsWorkbook = new XSSFWorkbook(file)) {

            Sheet urlsSheet = urlsWorkbook.getSheet("Mar1");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

            // Extract URLs from Excel
            for (int i = 1; i < rowCount; i++) { // Start from 1 to skip header
                Row row = urlsSheet.getRow(i);
                if (row.getCell(5) != null) {
                    String url;
                    if (row.getCell(5).getCellType() == CellType.STRING) {
                        url = row.getCell(5).getStringCellValue();
                    } else if (row.getCell(5).getCellType() == CellType.NUMERIC) {
                        url = String.valueOf(row.getCell(5).getNumericCellValue());
                    } else {
                        url = "NA"; // Handle cases where the cell is neither numeric nor string
                    }
                    
                    inputPid.add(row.getCell(0).getStringCellValue());
                    InputCity.add(row.getCell(1).getStringCellValue());
                    InputName.add(row.getCell(2).getStringCellValue());
                    InputSize.add(row.getCell(3).getStringCellValue());
                    NewProductCode.add(row.getCell(4).getStringCellValue());
                    uRL.add(url);
                    UOM.add(row.getCell(6).getStringCellValue());
                    Mulitiplier.add(row.getCell(7).getStringCellValue());
                   // Availability.add(row.getCell(8).getStringCellValue());
                    Pincode.add(row.getCell(9).getStringCellValue());
                    NameForCheck.add(row.getCell(10).getStringCellValue());
                }
            }

            Set<String> usedPincodes = new HashSet<>();
            

            // Create Excel workbook for storing results
            Workbook resultsWorkbook = new XSSFWorkbook();
            Sheet resultsSheet = resultsWorkbook.createSheet("Results");
            createHeaderRow(resultsSheet);

            int rowIndex = 1;
            for (int i = 0; i < uRL.size(); i++) {
                String url = uRL.get(i);
                String currentPincode = Pincode.get(i); // Get the current pincode from the list
                if (url.isEmpty() || url.equalsIgnoreCase("NA")) {
                    writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
                            InputSize.get(i), "NA", url, "NA", "NA", "NA", 
                            "NA", Mulitiplier.get(i), "NA", "NA", "NA", "NA", "NA", "NA", NameForCheck.get(i));
                    System.out.println("Skipped processing for URL: " + url);
                    continue;
                }

                try {
                    driver.get(url); 
                    options.addArguments("--auto-open-devtools-for-tabs");
                    Thread.sleep(2000);
//                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[4]/div"))).click();//div[@class='css-0']//div[@class='chakra-stack css-hwa0xx']
//                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[9]/div[3]/div/section/header/div/div[2]/div/input"))).sendKeys(Keys.ENTER);
//                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[9]/div[3]/div/section/header/div/div[2]/div/input"))).sendKeys("560001");
//                    Thread.sleep(2000); // Wait for the location dropdown to load
//
//                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[9]/div[3]/div/section/div/div[1]/div[1]"))).click();
//                    Thread.sleep(5000); // Wait for the page to load
                    
                   // wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[4]/div"))).click(); 
                   // wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[9]/div[3]/div/section/header/div/div[2]/div/input"))).sendKeys(Keys.ENTER);

                    // Check if current pincode has been used
                    if (!usedPincodes.contains(currentPincode)) {
                    	 wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[4]/div"))).click(); 
                         wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[9]/div[3]/div/section/header/div/div[2]/div/input"))).sendKeys(Keys.ENTER);
                        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[9]/div[3]/div/section/header/div/div[2]/div/input"))).sendKeys(currentPincode);
                        usedPincodes.add(currentPincode); // Add pincode to used set
                        wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[9]/div[3]/div/section/div/div[1]/div[1]"))).click();
                        Thread.sleep(2000);
                      //  Thread.sleep(5000); // Wait for the page to load
                        
                    } else {
                        System.out.println("Pincode " + currentPincode + " has already been used.");
                        // Optionally, you can skip sending the pincode input or handle accordingly
                    }
                    
                    System.out.println("PRODUCTCOUNT = " + ProductCOUNT);
                    
                     // Wait for the page to load
                    String productId = extractProductId(url);
                    System.out.println("Extracted Product ID: " + productId);
                    String newName = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='chakra-stack css-1u0scjk']//h1"))).getText();
                    System.out.println(newName);
                    //uom 
                    String weight = extractWeight(newName);
                    System.out.println("Extracted Weight: " + weight);
                    //String mrpValue = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='chakra-stack css-1k4nord']//td[@class='css-rq808y']"))).getText();
                    String mrpValue = extractMRP(driver);
                    System.out.println(mrpValue);
                    //String spValue =wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='chakra-stack css-1k4nord']//td[@class='css-1z07v0v']"))).getText();
                    String spValue = extractSP(driver);
                    System.out.println(spValue);
                   // String offerValue = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='chakra-stack css-1k4nord']//td[@class='css-1ocm65q'][2]"))).getText();
                    String offerValue = extractOffer(driver);
                    System.out.println(offerValue);
                   
                    
                    int result;
//                    if (url.contains("NA")) {
//                    	Availability = "NA";
//                    } else {
                        try {
                            // Define the texts to check for
                            String[] textsToCheck = {
                                "Not available",
                                "Out of Stock"
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
                       String availability =String.valueOf(result);
                       
                       ProductCOUNT++;
                 //   }

                    // Write results to the results sheet
                    writeResults(resultsSheet, rowIndex++, inputPid.get(i), InputCity.get(i), InputName.get(i),
                            InputSize.get(i), productId, url, newName, mrpValue, spValue, 
                            weight,String.valueOf(Mulitiplier.get(i)), availability, offerValue, "NA", "NA", "NA", "NA", NameForCheck.get(i));

                    System.out.println("Data extracted for URL: " + url);
                } catch (Exception e) {
                    e.printStackTrace();
                    System.out.println("Failed to extract data for URL: " + url);
                }
            }
            
            // Write results to Excel file
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());
            String outputFilePath = ".\\Output\\BIGBASKET_OutputData_" + timestamp + ".xlsx";
            
            // Write results to Excel file
            FileOutputStream outFile = new FileOutputStream(outputFilePath);
            resultsWorkbook.write(outFile);
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

    private static void createHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        String[] headers = {"InputPid", "InputCity", "InputName", "InputSize", "NewProductCode", "URL", "Name", "MRP", 
                            "SP", "UOM", "Multiplier", "Availability",  "Offer", "Commands", "Remarks", "Correctness", 
                            "Percentage", "Name", "NameForCheck"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }
    }

    private static void writeResults(Sheet sheet, int rowIndex, String id, String city, String name, String size,
                                      String productId, String url, String newName, String mrpValue,
                                      String spValue, String weight, String mulitiplier, String availability,String offerValue,
                                      String command, String remark, String correctness, String percentage,
                                       String nameCheck) {
        Row resultRow = sheet.createRow(rowIndex);
        resultRow.createCell(0).setCellValue(id);
        resultRow.createCell(1).setCellValue(city);
        resultRow.createCell(2).setCellValue(name);
        resultRow.createCell(3).setCellValue(size);
        resultRow.createCell(4).setCellValue(productId);
        resultRow.createCell(5).setCellValue(url);
        resultRow.createCell(6).setCellValue(newName);
        resultRow.createCell(7).setCellValue(mrpValue);
        resultRow.createCell(8).setCellValue(spValue);
        resultRow.createCell(9).setCellValue(weight);
        resultRow.createCell(10).setCellValue(mulitiplier);
        resultRow.createCell(11).setCellValue(availability);
        resultRow.createCell(12).setCellValue(offerValue);
        resultRow.createCell(13).setCellValue(command);
        resultRow.createCell(14).setCellValue(remark);
        resultRow.createCell(15).setCellValue(correctness);
        resultRow.createCell(16).setCellValue(percentage);
        resultRow.createCell(17).setCellValue(nameCheck);
        
    }

    private static String extractMRP(WebDriver driver) {
        try {
            String text = driver.findElement(By.xpath("//div[@class='chakra-stack css-1k4nord']//td[@class='css-rq808y']")).getText();
            return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
        } catch (Exception e) {
            e.printStackTrace();
            return "NA";
        }
    }

    private static String extractSP(WebDriver driver) {
        try {
            String text = driver.findElement(By.xpath("//div[@class='chakra-stack css-1k4nord']//td[@class='css-1z07v0v']")).getText();
            return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
        } catch (Exception e) {
            e.printStackTrace();
            return "NA";
        }
    }

    private static String extractOffer(WebDriver driver) {
        try {
            String text = driver.findElement(By.xpath("//div[@class='chakra-stack css-1k4nord']//td[@class='css-1ocm65q'][2]")).getText();
            //return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
            return isValidValue(text) ? text.trim() : "NA";
        } catch (Exception e) {
            e.printStackTrace();
            return "NA";
        }
    }

    private static String extractProductId(String url) {
        // Find the index of "/pd/"
        int startIndex = url.indexOf("/pd/") + 4; // Add 4 to move past "/pd/"
        
        // Find the index of the next "/"
        int endIndex = url.indexOf("/", startIndex);
        
        // If no further slash is found, extract till the end of the string
        if (endIndex == -1) {
            endIndex = url.length();
        }
        
        // Extract and return the product ID using substring
        return url.substring(startIndex, endIndex);
    }

    
    private static String extractWeight(String text) {
        // Find the last comma
        int lastCommaIndex = text.lastIndexOf(',');
        
        // Check if a comma is found
        if (lastCommaIndex != -1) {
            // Extract the substring after the last comma
            String weightInfo = text.substring(lastCommaIndex + 1).trim(); // trim to remove leading spaces
            
            // Since we expect 'g' to be at the end of the weight, we can return only the weight part.
            return weightInfo.split(" ")[0] + " " + weightInfo.split(" ")[1]; // Return the first two parts
        }
        
        return "Weight not found"; // Return a default message if no comma is found
    }


    private static boolean isValidValue(String value) {
        return value != null && !value.isEmpty() && !value.equals("₹");
    }

}
