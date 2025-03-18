package Shopping;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import org.openqa.selenium.TimeoutException;

public class BigBasketProductScraper {

    public static void main(String[] args) {
        // Generate timestamp
        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String timestamp = currentTime.format(formatter);

        String inputFile = ".\\input-data\\Product Search.xlsx";
        String outputFile = ".\\Output\\Bigbasket_Pro_Ser_OutputData " + timestamp + " .xlsx";

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
            headerRow.createCell(0).setCellValue("PID");
            headerRow.createCell(1).setCellValue("CITY");
            headerRow.createCell(2).setCellValue("Input Product Name");
            headerRow.createCell(3).setCellValue("UOM");
            headerRow.createCell(4).setCellValue("Product URL");
            headerRow.createCell(5).setCellValue("Product Name");
            headerRow.createCell(6).setCellValue("MRP");
            headerRow.createCell(7).setCellValue("SP");
            headerRow.createCell(8).setCellValue("uom");

            // Read input Excel sheet
            Workbook workbook = WorkbookFactory.create(new File(inputFile));
            Sheet sheet = workbook.getSheetAt(0); // Assuming input is in first sheet

            for (Row row : sheet) {
                if (row == null || row.getCell(0) == null || row.getCell(0).getStringCellValue().isEmpty()) {
                    break;
                }

                String pId = row.getCell(0).getStringCellValue();
                String city = row.getCell(1).getStringCellValue();
                String productName = row.getCell(2).getStringCellValue();
                String uom = row.getCell(3).getStringCellValue();

                driver.get("https://www.bigbasket.com/");

                // Find the search bar and search for each product
                WebElement searchInput = driver.findElement(By.xpath("/html/body/div[2]/div[1]/header[2]/div[1]/div[1]/div/div/div/div/input"));
                searchInput.clear(); // Clear the search bar
                searchInput.sendKeys(productName); // Enter the product name
                searchInput.sendKeys(Keys.ENTER); // Submit the search query

                // Wait for the search results to load
                try {
                    WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(20));
                    WebElement searchResultsSection = wait1.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//section[contains(@class, 'z-10')]")));
                    List<WebElement> productElements = searchResultsSection.findElements(By.tagName("li"));

                    // Initialize variables with "NA"
                    String productUrl = "NA";
                    String productNameResult = "NA";
                    String originalMrp = "NA";
                    String spValue = "NA";
                    String originalUom = "NA";

                    // Check if there are any products found
                    if (productElements.size() > 0) {
                        // Limit to scrape only the first three products
                        int count = Math.min(productElements.size(), 3);

                        // Loop through the first three product elements
                        for (int i = 0; i < count; i++) {
                            // Get the ith product element
                            WebElement productElement = productElements.get(i);

                            Thread.sleep(2000);

                            try {
                                // Find the anchor tag within the product element
                                WebElement productLink = productElement.findElement(By.xpath(".//a[@class='h-full']"));
                                productUrl = productLink.getAttribute("href");
                            } catch (Exception w) {
                            }

                            try {
                                WebElement productNewName = productElement.findElement(By.xpath(" .//h3[@class = 'block m-0 line-clamp-2 font-regular text-base leading-sm text-darkOnyx-800 pt-0.5 h-full']"));
                                productNameResult = productNewName.getText();
                            } catch (Exception e) {
                            }

                            try {
                                WebElement mrp = productElement.findElement(By.xpath(" .//span[@class = 'Label-sc-15v1nk5-0 Pricing___StyledLabel2-sc-pldi2d-2 gJxZPQ hsCgvu']"));
                                String originalMrp1 = mrp.getText();
                                originalMrp = originalMrp1.replace("₹", "");
                            } catch (Exception h) {
                            }

                            try {
                                WebElement sp = productElement.findElement(By.xpath(" .//span[@class = 'Label-sc-15v1nk5-0 Pricing___StyledLabel-sc-pldi2d-1 gJxZPQ AypOi']"));
                                String originalSp = sp.getText();
                                spValue = originalSp.replace("₹", "");
                            } catch (Exception t) {
                            }

                            try {
                                WebElement uomElement = productElement.findElement(By.xpath(".//span[@class = 'Label-sc-15v1nk5-0 gJxZPQ truncate']"));
                                originalUom = uomElement.getText();
                            } catch (Exception j) {
                                // Handle exception
                            }

                            // Write data to output Excel sheet
                            Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                            outputRow.createCell(0).setCellValue(pId);
                            outputRow.createCell(1).setCellValue(city);
                            outputRow.createCell(2).setCellValue(productName);
                            outputRow.createCell(3).setCellValue(uom);
                            outputRow.createCell(4).setCellValue(productUrl);
                            outputRow.createCell(5).setCellValue(productNameResult);
                            outputRow.createCell(6).setCellValue(originalMrp);
                            outputRow.createCell(7).setCellValue(spValue);
                            outputRow.createCell(8).setCellValue(originalUom);

                            System.out.println(productUrl);
                            System.out.println(productNameResult);
                            System.out.println(originalMrp);
                            System.out.println(spValue);
                            System.out.println(originalUom);
                        }
                    } else {
                        // No products found, write "NA" to the output Excel sheet for all fields
                        Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);
                        outputRow.createCell(0).setCellValue(pId);
                        outputRow.createCell(1).setCellValue(city);
                        outputRow.createCell(2).setCellValue(productName);
                        outputRow.createCell(3).setCellValue(uom);
                        outputRow.createCell(4).setCellValue(productUrl);
                        outputRow.createCell(5).setCellValue(productNameResult);
                        outputRow.createCell(6).setCellValue(originalMrp);
                        outputRow.createCell(7).setCellValue(spValue);
                        outputRow.createCell(8).setCellValue(originalUom);
                    }
                } catch (TimeoutException e) {
                    System.out.println("TimeoutException occurred. Skipping product: " + productName);
                    // Continue with the next product
                    continue;
                }
            }

            // Save output Excel file
            FileOutputStream fileOut = new FileOutputStream(outputFile);
            outputWorkbook.write(fileOut);
            fileOut.close();

            // Close the WebDriver
            System.out.println("DONE SCRAPING");
            driver.quit();

        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
}
