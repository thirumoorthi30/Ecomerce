package Shopping;

import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.time.Duration;
import java.util.Arrays;
import java.util.List;

public class PurplleScraper {
    public static void main(String[] args) throws InterruptedException {
        // Set the path to the ChromeDriver executable
        System.setProperty("webdriver.chrome.driver", "./Drivers//chromedriver.exe");

        // Initialize the Chrome driver with headless option
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--start-maximized");
        WebDriver driver = new ChromeDriver(options);

        // List to store product names
        List<String> products = Arrays.asList(
            "Boroplus Doodh Kesar Body Lotion (400 ml)",
            "Cetaphil Daily Advance Ultra Hydrating Moisturizing Lotion (Dermatologists Recommended)"
        );

        // Navigate to Purplle
        driver.get("https://www.purplle.com/");

        // Wait for search box to be visible
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
       // WebElement searchBox = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("input#search")));
        
        
        Thread.sleep(1000);
        // Find the search bar and search for each product
        WebElement searchInput = driver.findElement(By.xpath("/html/body/app-root/div/div/app-header/header/div[2]/div/div/div/div[1]/div/input"));

        for (int i = 0; i < Math.min(2, products.size()); i++) {
            String product = products.get(i);
            // Clear the search box and input the product name
            searchInput.click();
            
            
            searchInput.clear();
            WebElement Search=driver.findElement(By.xpath("/html/body/app-root/div/div/app-header/div[2]/desktop-search-content/div[1]/div/div[1]/p/input"));
            Search.sendKeys(product);
            Search.sendKeys(Keys.ENTER);
            
         // Wait for search results to be visible
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.d-block.w-100.position-relative.clearfix")));

            // Find the top three search results
            List<WebElement> searchResults = driver.findElements(By.cssSelector("div.d-block.w-100.position-relative.clearfix > app-listing-item"));
            
            
            System.out.println(searchResults);


            // Extract details from the top three search results
            for (int j = 0; j < Math.min(3, searchResults.size()); j++) {
                WebElement searchResult = searchResults.get(j);
                String productName = searchResult.findElement(By.xpath(".//div[@class = 'product-title fs-7 text-start text-black fw-normal']")).getText();
                String mrp = searchResult.findElement(By.xpath(".//s[@class = 'text-black-50 ms-1 fw-medium ng-star-inserted']")).getText();
                String sellingPrice = searchResult.findElement(By.xpath(".//span[@class = 'text-black fw-bolder fs-6']")).getText();
                System.out.println("Product Name: " + productName);
                System.out.println("MRP: " + mrp);
                System.out.println("Selling Price: " + sellingPrice);
                System.out.println("-------------------------");
            }
        }

        // Close the browser after scraping the products
        driver.quit();
    }
}
