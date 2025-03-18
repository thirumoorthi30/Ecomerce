package Shopping;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class FlipkartProductSearch {
    public static void main(String[] args) throws InterruptedException,WebDriverException {
        LocalDateTime currentTime = LocalDateTime.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        String timestamp = currentTime.format(formatter);

        String inputFile = ".\\input-data\\Product Search.xlsx";
        String outputFile = ".\\Output\\Amazon_Product_Search_OutputData_" + timestamp + ".xlsx";

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
            headerRow.createCell(0).setCellValue("Input Product ID");
            headerRow.createCell(1).setCellValue("Input Product Name");
            headerRow.createCell(2).setCellValue("Input Product Uom");
            headerRow.createCell(3).setCellValue("Product URL");
            headerRow.createCell(4).setCellValue("Product Name");
            headerRow.createCell(5).setCellValue("MRP");
            headerRow.createCell(6).setCellValue("SP");

            // Read input Excel sheet
            Workbook workbook = WorkbookFactory.create(new File(inputFile));
            Sheet sheet = workbook.getSheetAt(0); // Assuming input is in first sheet

            for (Row row : sheet) {
                if (row == null || row.getCell(0) == null || row.getCell(0).getStringCellValue().isEmpty()) {
                    break;
                }
                String iPid = row.getCell(0).getStringCellValue();
                String productName = row.getCell(1).getStringCellValue();
                String uom = row.getCell(2).getStringCellValue();

                // Assuming product name is in first column

                driver.get("https://www.flipkart.com/");

                // Find the search bar and search for each product
                WebElement searchInput = driver.findElement(By.xpath("//input[@class='Pke_EE']"));
                searchInput.sendKeys(productName);
                searchInput.sendKeys(Keys.ENTER); // Submit the search query

                // Find all product elements in the search results
                List<WebElement> productElements = driver.findElements(By.cssSelector("div._1AtVbE"));

                int count = 0;

                // Loop through the product elements and select non-sponsored products
             // Loop through the product elements and select non-sponsored products
                for (WebElement productElement : productElements) {
                    try {
                        // Check if the product is sponsored
                        productElement.findElement(By.xpath(".//div[@class = '_4HTuuX']"));
                        continue; // Skip sponsored products
                    } catch (NoSuchElementException ignored) {
                        // Product is not sponsored
                        count++;
                        try {
                            // Get necessary data for the product
                            WebElement productNameElement = productElement.findElement(By.xpath(".//a[@class = 's1Q9rs']"));
                            String productNameText = productNameElement.getText();

                            WebElement productLink = productElement.findElement(By.xpath(".//a[@class = 's1Q9rs']"));
                            String productUrl = productLink.getAttribute("href");

                            WebElement mrp = productElement.findElement(By.xpath(".//div[@class = '_3I9_wc']"));
                            String originalMrp = mrp.getText().replace("₹", "");

                            WebElement sp = productElement.findElement(By.xpath(".//div[@class = '_30jeq3']"));
                            String originalSp = sp.getText().replace("₹", "");

                            // Write data to output Excel sheet
                            Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum() + 1);

                            outputRow.createCell(0).setCellValue(iPid);
                            outputRow.createCell(1).setCellValue(productName);
                            outputRow.createCell(2).setCellValue(uom);
                            outputRow.createCell(3).setCellValue(productUrl);
                            outputRow.createCell(4).setCellValue(productNameText);
                            outputRow.createCell(5).setCellValue(originalMrp);
                            outputRow.createCell(6).setCellValue(originalSp);

                            // Print product details
                            System.out.println("Product Name: " + productNameText);
                            System.out.println("Product URL: " + productUrl);
                            System.out.println("MRP: " + originalMrp);
                            System.out.println("SP: " + originalSp);

                            if (count >= 5) {
                                break; // Found five non-sponsored products, exit loop
                            }
                        } catch (NoSuchElementException e) {
                            // handle exception
                        }
                    }
                }

                // Save output Excel file
                FileOutputStream fileOut = new FileOutputStream(outputFile);
                outputWorkbook.write(fileOut);
                fileOut.close();

                

                // Close the WebDriver
                //driver.quit();
            }
            System.out.println("Done Scraping");
            driver.quit();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
