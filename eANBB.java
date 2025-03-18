package Dailyrun;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeOptions;
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
        ChromeOptions options = new ChromeOptions();
      //  EdgeOptions options = new EdgeOptions();
        options.addArguments("--start-maximized");
        options.addArguments("--window-size=375,812");
        options.addArguments("--user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 13_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0 Mobile/15E148 Safari/604.1");
      //  options.addArguments("--incognito");
        
        WebDriver driver = new ChromeDriver(options);
       // WebDriver driver = new EdgeDriver(options);
     //   driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        List<String> eanNumber = new ArrayList<>();
        List<String> inputID = new ArrayList<>();
        Workbook resultsWorkbook = new XSSFWorkbook();
        Sheet resultsSheet = resultsWorkbook.createSheet("Results");
        createHeaderRow(resultsSheet);

        int rowIndex = 1;

        try (FileInputStream file = new FileInputStream("./input-data/BB28feb.xlsx");
             Workbook urlsWorkbook = new XSSFWorkbook(file)) {

            Sheet urlsSheet = urlsWorkbook.getSheet("Sheet1");
            int rowCount = urlsSheet.getPhysicalNumberOfRows();

            // Extract EAN Numbers and Input ID from Excel
            for (int i = 1; i < rowCount; i++) { // Start from 1 to skip header
                Row row = urlsSheet.getRow(i);
                if (row.getCell(0) != null) {
                    eanNumber.add(row.getCell(0).getStringCellValue());
                    inputID.add(row.getCell(1).getStringCellValue());
                }
            }

            // Process each EAN number and inputID
            for (int i = 0; i < eanNumber.size(); i++) {
                try {
                    driver.get("https://www.google.com");
                    WebElement searchBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("q")));
                    String searchTerm = eanNumber.get(i) + " in jiomart";
                    searchBox.sendKeys(searchTerm);
                    searchBox.sendKeys(Keys.RETURN);

                    if(i == 0) {
                        Thread.sleep(30000);
                    }

                    WebElement firstResultLink = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("h3")));
                    firstResultLink.click();

                    wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//h1"))); 

                    String url = driver.getCurrentUrl();
                    String productId = extractProductId(url);
                    String newName = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@id='pdp_product_name']"))).getText();
                    String weight = "NA";
                    String spValue = extractSP(driver);
                    String mrpValue = extractMRP(driver);
                    String offerValue = extractOffer(driver);

                    // Write results to the results sheet, including inputID
                    writeResults(resultsSheet, rowIndex++, eanNumber.get(i), inputID.get(i), url, productId, newName, weight, mrpValue, spValue, offerValue, "");

                } catch (Exception e) {
                    e.printStackTrace();
                    System.out.println("Failed to extract data for search term: " + eanNumber.get(i));
                    writeResults(resultsSheet, rowIndex++, eanNumber.get(i), inputID.get(i), "NA", "NA", "NA", "NA", "NA", "NA", "NA", "Error retrieving data for EAN: " + eanNumber.get(i));
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("An error occurred during the extraction process.");
        } finally {
            saveResultsToExcel(resultsWorkbook);

            if (driver != null) {
                System.out.println("Closing the driver.");
                driver.quit();
            }
        }
    }

    private static void createHeaderRow(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        String[] headers = {"EAN Number", "INPUTID", "Url", "Product ID", "Name", "Weight", "MRP", "SP", "Offer", "Error Message"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }
    }

    private static void writeResults(Sheet sheet, int rowIndex, String eanNumber, String inputID, String url, String productId,
            String newName, String weight, String mrpValue, String spValue, String offerValue, String errorMessage) {
        Row resultRow = sheet.createRow(rowIndex);
        resultRow.createCell(0).setCellValue(eanNumber);
        resultRow.createCell(1).setCellValue(inputID);
        resultRow.createCell(2).setCellValue(url);
        resultRow.createCell(3).setCellValue(productId);
        resultRow.createCell(4).setCellValue(newName);
        resultRow.createCell(5).setCellValue(weight);
        resultRow.createCell(6).setCellValue(mrpValue);
        resultRow.createCell(7).setCellValue(spValue);
        resultRow.createCell(8).setCellValue(offerValue);
        resultRow.createCell(9).setCellValue(errorMessage);
    }

    private static void saveResultsToExcel(Workbook resultsWorkbook) {
        try {
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
            String timestamp = dateFormat.format(new Date());
            String outputFilePath = "./Output/JIOMART_OutputData_" + timestamp + ".xlsx";

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
            List<WebElement> elements = driver.findElements(By.xpath("//div[contains(@class, 'jm-body-s jm-fc-primary-grey-80')]//span[contains(text(), '₹')]"));
            if (elements.isEmpty()) {
                return "NA";  // MRP value not found
            }
            String text = elements.get(0).getText();
            return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
        } catch (Exception e) {
            return "NA";  // In case of any other exception
        }
    }

    private static String extractSP(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("//span[contains(@class, 'jm-heading-xs jm-ml-xxs')]"));
            if(elements.isEmpty()) {
                return "NA";
            }
            String text = elements.get(0).getText();
            return isValidValue(text) ? text.replace("₹", "").trim() : "NA";
        } catch (Exception e) {
            e.printStackTrace();
            return "NA";
        }
    }

    private static String extractOffer(WebDriver driver) {
        try {
            List<WebElement> elements = driver.findElements(By.xpath("//span[contains(@class, 'jm-badge jm-ml-base jm-body-s-link')]"));
            if(elements.isEmpty()) {
                return "NA";
            }
            String text = elements.get(0).getText();
            return isValidValue(text) ? text.trim() : "NA";
        } catch (Exception e) {
            e.printStackTrace();
            return "NA";
        }
    }

    private static String extractProductId(String url) {
        int lastSlashIndex = url.lastIndexOf("/") + 1;
        return url.substring(lastSlashIndex);
    }

    private static boolean isValidValue(String value) {
        return value != null && !value.isEmpty() && !value.equals("₹");
    }
}
