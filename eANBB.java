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

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class eANBB {
    public static void main(String[] args) throws Exception {
        // System.setProperty("webdriver.chrome.driver", "./Drivers/chromedriver.exe");

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");
        options.addArguments("--window-size=375,812");
        options.addArguments("--user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 13_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0 Mobile/15E148 Safari/604.1");

        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        List<String> eanNumber = new ArrayList<>();
        Workbook resultsWorkbook = new XSSFWorkbook();
        Sheet resultsSheet = resultsWorkbook.createSheet("Results");
        createHeaderRow(resultsSheet);

        int rowIndex = 1;

        try (FileInputStream file = new FileInputStream("./input-data/BB28feb.xlsx");
             Workbook urlsWorkbook = new XSSFWorkbook(file)) {

            Sheet urlsSheet = urlsWorkbook.getSheet("Sheet1");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

            // Extract EAN Numbers from Excel
            for (int i = 1; i < rowCount; i++) { // Start from 1 to skip header
                Row row = urlsSheet.getRow(i);
                if (row.getCell(0) != null) {
                    eanNumber.add(row.getCell(0).getStringCellValue());
                }
            }

            for (int i = 0; i < eanNumber.size(); i++) { 
                try {
                	
                    // Open Google
                    driver.get("https://www.google.com");
                    driver.manage().deleteAllCookies();
              
                    WebElement searchBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("q")));

                    // Enter search terms
                    String searchTerm = eanNumber.get(i) + " in bigbasket"; 
                    searchBox.sendKeys(searchTerm);
                    searchBox.sendKeys(Keys.RETURN);

                    // Click on the first link in the search results
                    WebElement firstResultLink = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("h3")));
                    firstResultLink.click();

                    // Wait for the new page to load
                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1"))); 
                    

                    // Extract information
                    
                    String url = driver.getCurrentUrl();
                    System.out.println(url);
                    
                    String productId = extractProductId(url);
                    System.out.println(productId);
                    String newName = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='chakra-stack css-1u0scjk']//h1"))).getText();
                    System.out.println(newName);
                    String weight = extractWeight(newName);
                    System.out.println(weight);
                    String mrpValue = extractMRP(driver);
                    System.out.println(mrpValue);
                    String spValue = extractSP(driver);
                    System.out.println(spValue);
                    String offerValue = extractOffer(driver);
                    System.out.println(offerValue);
                    
                    // Write results to the results sheet
                    writeResults(resultsSheet, rowIndex++, eanNumber.get(i), url, productId, newName, weight, mrpValue, spValue, offerValue);

                } catch (Exception e) {
                    e.printStackTrace();
                    System.out.println("Failed to extract data for search term: " + eanNumber.get(i));
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
        String[] headers = {"EAN Number", "Url", "Product ID", "Name", "Weight", "MRP", "SP", "Offer"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }
    }

    private static void writeResults(Sheet sheet, int rowIndex, String eanNumber, String url, String productId,
            String newName, String weight, String mrpValue, String spValue, String offerValue) {
        Row resultRow = sheet.createRow(rowIndex);
        resultRow.createCell(0).setCellValue(eanNumber);
        resultRow.createCell(1).setCellValue(url);
        resultRow.createCell(2).setCellValue(productId);
        resultRow.createCell(3).setCellValue(newName);
        resultRow.createCell(4).setCellValue(weight);
        resultRow.createCell(5).setCellValue(mrpValue);
        resultRow.createCell(6).setCellValue(spValue);
        resultRow.createCell(7).setCellValue(offerValue);
    }

    private static void saveResultsToExcel(Workbook resultsWorkbook) {
        try {
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());
            String outputFilePath = "./Output/BIGBASKET_OutputData_" + timestamp + ".xlsx";

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
            return isValidValue(text) ? text.trim() : "NA";
        } catch (Exception e) {
            e.printStackTrace();
            return "NA";
        }
    }

    private static String extractProductId(String url) {
        int startIndex = url.indexOf("/pd/") + 4;
        int endIndex = url.indexOf("/", startIndex);
        if (endIndex == -1) {
            endIndex = url.length();
        }
        return url.substring(startIndex, endIndex);
    }

    private static String extractWeight(String text) {
        int lastCommaIndex = text.lastIndexOf(',');
        if (lastCommaIndex != -1) {
            String weightInfo = text.substring(lastCommaIndex + 1).trim();
            return weightInfo.split(" ")[0] + " " + weightInfo.split(" ")[1];
        }
        return "Weight not found";
    }

    private static boolean isValidValue(String value) {
        return value != null && !value.isEmpty() && !value.equals("₹");
    }
}
