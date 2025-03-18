package Shopping;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class resSwiggyfinal {
    public static void main(String[] args) {
        WebDriver driver = new ChromeDriver();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        try {
            // Load input data
            FileInputStream file = new FileInputStream(new File(".\\input-data\\resSwiggy.xlsx"));
            Workbook inputWorkbook = new XSSFWorkbook(file);
            Sheet inputSheet = inputWorkbook.getSheetAt(0);
            List<String> productNames = new ArrayList<>();
            List<String> pinCodes = new ArrayList<>();

            // Collect product names and pin codes
            Iterator<Row> iterator = inputSheet.iterator();
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Cell dishCell = currentRow.getCell(0);
                Cell pinCell = currentRow.getCell(1);
                if (dishCell != null && dishCell.getCellType() != CellType.BLANK) {
                    productNames.add(dishCell.getStringCellValue());
                }
                if (pinCell != null && pinCell.getCellType() != CellType.BLANK) {
                    String pinCode = pinCell.getStringCellValue();
                    if (!pinCodes.contains(pinCode)) {
                        pinCodes.add(pinCode);
                    }
                }
            }

            // Iterate over each pin code
            for (String locationPin : pinCodes) {
                // Create a new workbook for each pin code
                Workbook outputWorkbook = new XSSFWorkbook();
                Sheet outputSheet = outputWorkbook.createSheet("Swiggy Data");

                // Set up header row for the output file
                Row headerRow = outputSheet.createRow(0);
                headerRow.createCell(0).setCellValue("Product Name");
                headerRow.createCell(1).setCellValue("Restaurant Name");
                headerRow.createCell(2).setCellValue("Rating");
                headerRow.createCell(3).setCellValue("Delivery Time");
                headerRow.createCell(4).setCellValue("Dish Name");
                headerRow.createCell(5).setCellValue("Price");
                headerRow.createCell(6).setCellValue("Pin Code");

                // Open the browser and load the Swiggy site
                driver.get("https://www.swiggy.com/");
                driver.manage().window().maximize();

                // Set location pin
                WebElement location = driver.findElement(By.xpath("//*[@id=\"root\"]/div[1]/header/div/div/div/span[1]"));
                location.click();
                WebElement pinInput = wait.until(ExpectedConditions.visibilityOfElementLocated(
                        By.xpath("//*[@id=\"overlay-sidebar-root\"]/div/div/div[2]/div/div/div[2]/div[2]/div/input")));
                pinInput.clear();
                pinInput.sendKeys(locationPin);
                pinInput.sendKeys(Keys.ENTER);
                Thread.sleep(2000);
                wait.until(ExpectedConditions.visibilityOfElementLocated(
                        By.xpath("//*[@id=\"overlay-sidebar-root\"]/div/div/div[2]/div/div/div[3]/div/div/div[1]"))).click();
                Thread.sleep(1000);

                // Loop through all products for the current location
                for (String productName : productNames) {
                    try {
                        WebElement searchBar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div[1]/header/div/div/ul/li[5]/div/a")));
                        searchBar.click();
                    } catch (Exception e) {
                        try {
                            WebElement searchBar = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div[1]/div/div/div[2]/div[2]/div[2]/div/div[1]/div")));
                            searchBar.click();
                        } catch (Exception ee) {
                            // Handle error if search bar is not found
                        }
                    }

                    // Search for the product
                    WebElement searchInput = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div[1]/div/div[1]/div/form/div/div[1]/input")));
                    searchInput.click();
                    Thread.sleep(1000);
                    searchInput.clear();
                    Thread.sleep(2000);
                    searchInput.sendKeys(productName);
                    Thread.sleep(1000);
                    WebElement searchResult = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div[1]/div/div[2]/div/div/button[1]")));
                    Thread.sleep(1000);
                    searchResult.click();

                    Thread.sleep(2000);

                    WebElement dishTabClick = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id=\"root\"]/div[1]/div/div[2]/div/div/div[2]/div[1]/span[2]/span")));
                    dishTabClick.click();
                    Thread.sleep(1500);

                    // Scrape restaurant data
                    while (true) {
                        List<WebElement> restaurants = driver.findElements(By.xpath("//div[@class='Search_widgetsV2__27BBR Search_widgets__3o_bA']"));
                        System.out.println("Searching for: " + productName + " in PIN: " + locationPin);
                        System.out.println("Found restaurants: " + restaurants.size());

                        if (restaurants.isEmpty()) {
                            System.out.println("No restaurants found for: " + productName + " in PIN: " + locationPin);
                            continue; // Move to the next product
                        }

                        for (WebElement restaurant : restaurants) {
                            String restaurantName = restaurant.findElement(By.xpath(".//div[contains(@class, 'styles_restaurantName__5VIQZ')]")).getText();
                            String rating = restaurant.findElement(By.xpath(".//span[contains(@class, 'styles_restaurantMetaRating__3MhTg')]")).getText();
                            String timeForDelivery = restaurant.findElement(By.xpath(".//div[contains(@class, 'styles_restaurantMeta__2yx7V')]//div[2]")).getText();
                            String dishName = restaurant.findElement(By.xpath(".//div[contains(@class, 'sc-aXZVg cjJTeQ')]")).getText();
                            String price = restaurant.findElement(By.xpath(".//div[contains(@class, 'sc-aXZVg kCbDOU')]")).getText();

                            System.out.println("Restaurant: " + restaurantName);
                            System.out.println("Rating: " + rating);
                            System.out.println("Delivery Time: " + timeForDelivery);
                            System.out.println("Dish Name: " + dishName);
                            System.out.println("Price: " + price);
                            System.out.println("--------------------------------");

                            // Add data to the Excel file
                            Row dataRow = outputSheet.createRow(outputSheet.getPhysicalNumberOfRows());
                            dataRow.createCell(0).setCellValue(productName);
                            dataRow.createCell(1).setCellValue(restaurantName);
                            dataRow.createCell(2).setCellValue(rating);
                            dataRow.createCell(3).setCellValue(timeForDelivery);
                            dataRow.createCell(4).setCellValue(dishName);
                            dataRow.createCell(5).setCellValue(price);
                            dataRow.createCell(6).setCellValue(locationPin);
                        }

                        break;
                    }
                }

                // Save the output Excel file for the current pin code
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
                String timestamp = dateFormat.format(new Date());
                String outputFilePath = ".\\Output\\Swiggy_Restaurant_Output_" + locationPin + "_" + timestamp + ".xlsx";
                try (FileOutputStream outFile = new FileOutputStream(outputFilePath)) {
                    outputWorkbook.write(outFile);
                    System.out.println("Output file saved for pin " + locationPin + ": " + outputFilePath);
                } catch (IOException e) {
                    e.printStackTrace();
                } finally {
                    // Close the workbook after saving
                    outputWorkbook.close();
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }
}
