package DataAll;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotInteractableException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class swiggyAll {

    public static void main(String[] args) throws Exception {

        // Set up the WebDriver and Excel workbook
        System.setProperty("webdriver.chrome.driver", "./Drivers/chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        String url = "https://www.swiggy.com/instamart";
        driver.get(url);

        Thread.sleep(4000);

        // Click on the Fresh Vegetables category
        driver.findElement(By.xpath("//div[text()='Fresh Vegetables']")).click();

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(50));
        wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@class='sc-cwHptR jCovMf']//li")));
        List<WebElement> Subcatcount = driver.findElements(By.xpath("//div[@class='sc-cwHptR jCovMf']//li"));
        System.out.println("Number of subcategories: " + Subcatcount.size());

        // Create a new Excel Workbook and Sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Fresh Vegetables");

        // Create the header row
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Subcategory");
        header.createCell(1).setCellValue("Product Name");
        header.createCell(2).setCellValue("MRP");
        header.createCell(3).setCellValue("Selling Price");
        header.createCell(4).setCellValue("UOM");
        header.createCell(5).setCellValue("Offer");
        header.createCell(6).setCellValue("Availability");
        header.createCell(7).setCellValue("Page URL");  // Add the new column for URL

        int rowIndex = 1; // Start writing from the second row (after the header)

        // Check and create the 'Output' directory if it does not exist
        File outputDir = new File("./Output");
        if (!outputDir.exists()) {
            outputDir.mkdir();  // Create the directory if it does not exist
        }

        // Define the full path for the Excel file
        String outputPath = ".\\Output\\Swiggy_Fresh_Vegetables.xlsx";

        for (int i = 0; i < Subcatcount.size(); i++) {
            try {
                // Scroll to bottom and wait for new content
                Thread.sleep(5000); // Wait for the new content to load

                WebElement Subcategory = Subcatcount.get(i);
                String SubCatName = Subcategory.getText();
                System.out.println("Subcategory Name: " + SubCatName);

                // Wait for fresh vegetables list to load
                wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@class='K0-3A _1Nlcf']")));

                // Ensure the subcategory is visible and clickable
                scrollToElement(driver, Subcategory);  // Scroll to the subcategory
                Subcategory.click();  // Click the subcategory

                // Wait for fresh vegetables list to load
                wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[@class='K0-3A _1Nlcf']")));
                
                WebElement productListContainer = driver.findElement(By.xpath("//div[@class='Xxq5s']"));
                productListContainer.click();
                Thread.sleep(2000);

                // Scroll the container to the bottom to load all products
                simulateScroll(driver);

                // Grab all fresh vegetables listed in the subcategory
                List<WebElement> Freshvegcatcount = driver.findElements(By.xpath("//div[@class='K0-3A _1Nlcf']"));
                System.out.println("Number of fresh vegetables in subcategory: " + Freshvegcatcount.size());

                for (int j = 1; j<=Freshvegcatcount.size(); j++) {
                    try {
                        // Use XPath to locate the image element
                    	
                    	//WebDriverWait waitForLoaderToDisappear = new WebDriverWait(driver, Duration.ofSeconds(10));
                    	//itForLoaderToDisappear.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("(//img[@data-testid='item-image-default'])[" + j +"]"))); // Replace with actual loader element ID

                        WebElement inside = driver.findElement(By.xpath("(//img[@data-testid='item-image-default'])[" + j  + "]"));

                        // Wait for the element to be clickable (visible and interactable)
                        wait.until(ExpectedConditions.elementToBeClickable(inside));

                        // Scroll the element into view if needed
                        scrollToElement(driver, inside);

                        // Try clicking the element
                        //inside.click();
                        JavascriptExecutor js = (JavascriptExecutor) driver;
                        js.executeScript("arguments[0].click();", inside);

                    	
                        Thread.sleep(1000); // Wait for 1 second after clicking

                        // Extract product details
                        WebElement Name = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[@class='sc-aXZVg bzVIAg _AHZN']")));
                        String productName = Name.getText();

                        String stockStatus = "1";  // Default to '0' (sold out)
                        try {
                            // Try to find the availability element
                            driver.findElement(By.xpath(".//div[@class='sc-aXZVg lfvVbl _1Hegg']"));
                            stockStatus = "0";  // If the element is found, set availability to '1' (in stock)
                        } catch (Exception e) {
                            System.out.println("Product is sold out or does not have availability element.");
                        }

                        WebElement MRP = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[@class='sc-aXZVg fVWuLc _2XPBo _1QyO8']")));
                        String productMRP = MRP.getText();
                        WebElement SP = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[@class='sc-aXZVg bzVIAg _2XPBo']")));
                        String productSP = SP.getText();
                        WebElement UOM = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[@class='sc-aXZVg fVWuLc _1TwvP']")));
                        String productUOM = UOM.getText();
                        WebElement Offer = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(".//div[@class='sc-aXZVg csCGOB _1kaS2']")));
                        String productOffer = Offer.getText();
                        String pageurl = driver.getCurrentUrl(); // Get the current URL

                        // Output product details
                        System.out.println("Product Name: " + productName);
                        System.out.println("MRP: " + productMRP);
                        System.out.println("Selling Price: " + productSP);
                        System.out.println("UOM: " + productUOM);
                        System.out.println("Offer: " + productOffer);
                        System.out.println("Stock Availability: " + stockStatus);
                        System.out.println("URL: " + pageurl);

                        // Write data to Excel
                        Row row = sheet.createRow(rowIndex++);
                        row.createCell(0).setCellValue(SubCatName);  // Write Subcategory name here
                        row.createCell(1).setCellValue(productName);
                        row.createCell(2).setCellValue(productMRP);
                        row.createCell(3).setCellValue(productSP);
                        row.createCell(4).setCellValue(productUOM);
                        row.createCell(5).setCellValue(productOffer);
                        row.createCell(6).setCellValue(stockStatus);  // Write stock status (0 for "Sold Out", 1 for available)
                        row.createCell(7).setCellValue(pageurl);  // Write the page URL

                    } catch (Exception e) {
                        System.out.println("Error processing product at index " + j + ": " + e.getMessage());
                    } finally {
                        // Click the back button to go back to the product list
                        try {
                            driver.findElement(By.xpath(".//div[@class='_13vAM']")).click();
                        } catch (Exception e) {
                            System.out.println("Error navigating back: " + e.getMessage());
                        }
                    }
                }
            } catch (Exception e) {
                System.out.println("Error processing subcategory " + i + ": " + e.getMessage());
            }
        }

        // Write the data to an Excel file
        try (FileOutputStream fileOut = new FileOutputStream(new File(outputPath))) {
            workbook.write(fileOut);
            System.out.println("Excel file has been written successfully!");
        } catch (IOException e) {
            System.err.println("Error writing the Excel file: " + e.getMessage());
        }

        // Close the workbook
        workbook.close();

        driver.quit();
    }

    // Function to scroll to an element
    public static void scrollToElement(WebDriver driver, WebElement element) {
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("arguments[0].scrollIntoView(true);", element);
    }

    // Simulate scroll by sending 'END' key multiple times to load more products
    public static void simulateScroll(WebDriver driver) throws InterruptedException {
        Actions actions = new Actions(driver);
        for (int i = 0; i < 20; i++) {
            actions.sendKeys(Keys.END).perform();
            Thread.sleep(2000);  // Wait for content to load
        }
    }
}
