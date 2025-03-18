package Shopping;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.Assert;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class Extractproductmatch {

    public static void main(String[] args) {
        // Generate timestamp
        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String timestamp = currentTime.format(formatter);

        String inputFile = ".\\input-data\\Product Search.xlsx";
        String outputFile = ".\\Output\\Amazon_Product_Ser_OutputData_" + timestamp + ".xlsx";

        // Set Chrome driver path
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");

        // Initialize Chrome options
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");

        // Initialize WebDriver
        WebDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        try {
            // Initialize output Excel workbook
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("Results");

            Row headerRow = outputSheet.createRow(0);
            headerRow.createCell(0).setCellValue("Input Product Name");
            headerRow.createCell(1).setCellValue("Product URL");
            headerRow.createCell(2).setCellValue("Product Name");
            headerRow.createCell(3).setCellValue("MRP");
            headerRow.createCell(4).setCellValue("SP");

            // Read input Excel sheet
            Workbook workbook = WorkbookFactory.create(new File(inputFile));
            Sheet sheet = workbook.getSheetAt(0); // Assuming input is in first sheet

            for (Row row : sheet) {
                String productName = row.getCell(0).getStringCellValue(); // Assuming product name is in first column

                driver.get("https://www.amazon.in/");

                // Search for the specific product
                WebElement searchInput = driver.findElement(By.id("twotabsearchtextbox"));
                searchInput.sendKeys(productName);
                searchInput.sendKeys(Keys.RETURN);

                // Wait for the search results to load
                Thread.sleep(4000);

                // Find all product elements in the search results
                List<WebElement> productElements = driver.findElements(By.xpath("//div[@data-component-type='s-search-result']"));

                int count = 0;
                // Loop through the product elements and select non-sponsored products
                for (WebElement productElement : productElements) {
                    try {
                        // Check if the product is sponsored
                        productElement.findElement(By.className("puis-label-popover-default"));
                        continue; // Skip sponsored products
                    } catch (org.openqa.selenium.NoSuchElementException ignored) {
                        // Product is not sponsored
                        count++;

                        // Get necessary data for the product
                        WebElement productNameElement = productElement.findElement(By.xpath(".//h2/a/span"));
                        String productNameText = productNameElement.getText();

                        WebElement productLink = productElement.findElement(By.xpath(".//h2/a"));
                        String productUrl = productLink.getAttribute("href");

                        WebElement mrp = productElement.findElement(By.xpath(".//div[@class='a-section aok-inline-block']//span[@class='a-price a-text-price']//span[@aria-hidden]"));
                        String originalMrp = mrp.getText().replaceAll("[^\\d.]+", "");

                        WebElement sp = productElement.findElement(By.xpath(".//span[@class='a-price-whole']"));
                        String originalSp = sp.getText().replace("₹", "");

                        // Write data to output Excel sheet
                        Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                        outputRow.createCell(0).setCellValue(productName);
                        outputRow.createCell(1).setCellValue(productUrl);
                        outputRow.createCell(2).setCellValue(productNameText);
                        outputRow.createCell(3).setCellValue(originalMrp);
                        outputRow.createCell(4).setCellValue(originalSp);

                        System.out.println("Product Name: " + productNameText);
                        System.out.println("Product URL: " + productUrl);
                        System.out.println("MRP: " + originalMrp);
                        System.out.println("SP: " + originalSp);

                        if (count >= 3) {
                            break; // Found three non-sponsored products, exit loop
                        }
                    }
                    if(count >= 3) {
                    	break;
                    }
                }
            }

            // Save output Excel file
            FileOutputStream fileOut = new FileOutputStream(outputFile);
            outputWorkbook.write(fileOut);
            fileOut.close();

            // Close the WebDriver
            driver.quit();

        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
        System.out.println("Tsest");
        driver.quit();
        
    /*    finally {
            // Close the WebDriver in a finally block to ensure it's always closed
            if (driver != null) {
                driver.quit();
            }
        }   */
    }
}






